"""Чтение students.xlsx, парсинг tickets.docx и работа с журналом results.xlsx.

Модуль не зависит от GUI, поэтому его легко тестировать.
"""
from __future__ import annotations

import logging
import os
import random
import re
from dataclasses import dataclass
from datetime import datetime

from docx import Document
from openpyxl import Workbook, load_workbook

log = logging.getLogger("tickets")

RESULTS_HEADER = ["Группа", "Фамилия", "Имя", "Номер билета", "Дата и время", "Повтор (да/нет)"]
DATE_FMT = "%Y-%m-%d %H:%M:%S"
QUESTIONS_PER_TICKET = 3

_TICKET_RE = re.compile(r"^\s*Билет\s*№?\s*(\d+)\s*[.:]?\s*$", re.IGNORECASE)
_TICKET_LIKE_RE = re.compile(r"^\s*Билет\b", re.IGNORECASE)
_QUESTION_RE = re.compile(r"^\s*(\d+)\s*[.)]\s*(.+?)\s*$")


class SourceFileError(Exception):
    """Входной файл отсутствует или не читается."""


class ResultsLockedError(Exception):
    """results.xlsx занят другой программой (например, открыт в Excel)."""


@dataclass(frozen=True)
class Assignment:
    ticket_number: int
    is_repeat: bool


# ---------------------------------------------------------------- students.xlsx

def load_students(path: str) -> dict[str, list[tuple[str, str]]]:
    """Возвращает {имя_листа: [(Фамилия, Имя), ...]}. Строка 1 — заголовок."""
    if not os.path.isfile(path):
        raise SourceFileError(f"Не найден файл со студентами: {path}")
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:  # битый файл, не xlsx и т.п.
        raise SourceFileError(f"Не удалось прочитать {path}: {exc}") from exc

    groups: dict[str, list[tuple[str, str]]] = {}
    try:
        for ws in wb.worksheets:
            students = []
            for row in ws.iter_rows(min_row=2, max_col=2, values_only=True):
                surname = _cell_str(row[0] if len(row) > 0 else None)
                name = _cell_str(row[1] if len(row) > 1 else None)
                if surname or name:
                    students.append((surname, name))
            groups[ws.title] = students
    finally:
        wb.close()
    return groups


def students_of_group(groups: dict[str, list[tuple[str, str]]], group: str) -> list[str]:
    """Студенты группы в формате «Фамилия Имя». Неизвестная/пустая группа → []."""
    return [f"{s} {n}".strip() for s, n in groups.get(group, [])]


def _cell_str(value) -> str:
    return "" if value is None else str(value).strip()


# ---------------------------------------------------------------- tickets.docx

def _paragraph_is_numbered(paragraph) -> bool:
    """Автонумерация Word: у абзаца есть numPr или стиль списка."""
    ppr = paragraph._p.pPr
    if ppr is not None and ppr.numPr is not None:
        return True
    style = (paragraph.style.name or "").lower() if paragraph.style is not None else ""
    return "list number" in style or "список" in style


def parse_tickets_paragraphs(paragraphs: list[tuple[str, bool]]) -> dict[int, list[str]]:
    """Парсит список (текст, автонумерация?) в {номер_билета: [3 вопроса]}.

    Билет с нераспознанным заголовком или меньше чем 3 вопросами пропускается,
    причина пишется в лог.
    """
    tickets: dict[int, list[str]] = {}
    current: int | None = None
    questions: list[str] = []
    skipping = False  # находимся внутри блока с нераспознанным заголовком

    def flush():
        if current is None:
            return
        if len(questions) < QUESTIONS_PER_TICKET:
            log.error("Билет %s пропущен: вопросов %d, нужно %d",
                      current, len(questions), QUESTIONS_PER_TICKET)
        elif current in tickets:
            log.error("Билет %s пропущен: номер повторяется", current)
        else:
            if len(questions) > QUESTIONS_PER_TICKET:
                log.warning("Билет %s: вопросов %d, взяты первые %d",
                            current, len(questions), QUESTIONS_PER_TICKET)
            tickets[current] = questions[:QUESTIONS_PER_TICKET]

    for raw, auto_numbered in paragraphs:
        text = (raw or "").strip()
        if not text:
            continue
        m = _TICKET_RE.match(text)
        if m:
            flush()
            current, questions, skipping = int(m.group(1)), [], False
            continue
        if _TICKET_LIKE_RE.match(text):
            flush()
            log.error("Нераспознанный заголовок билета: %r — блок пропущен", text)
            current, questions, skipping = None, [], True
            continue
        if skipping:
            continue
        if current is None:
            log.warning("Текст вне билета проигнорирован: %r", text)
            continue
        q = _QUESTION_RE.match(text)
        if q:
            questions.append(q.group(2))
        elif auto_numbered:
            questions.append(text)
        else:
            log.warning("Билет %s: строка не похожа на вопрос: %r", current, text)
    flush()
    return tickets


def load_tickets(path: str) -> dict[int, list[str]]:
    if not os.path.isfile(path):
        raise SourceFileError(f"Не найден файл с билетами: {path}")
    try:
        doc = Document(path)
    except Exception as exc:
        raise SourceFileError(f"Не удалось прочитать {path}: {exc}") from exc
    paragraphs = [(p.text, _paragraph_is_numbered(p)) for p in doc.paragraphs]
    tickets = parse_tickets_paragraphs(paragraphs)
    log.info("Загружено билетов: %d", len(tickets))
    return tickets


# ---------------------------------------------------------------- results.xlsx

def ensure_results(path: str) -> None:
    """Создаёт results.xlsx с заголовком, если файла нет."""
    if os.path.isfile(path):
        return
    wb = Workbook()
    ws = wb.active
    ws.title = "Журнал"
    ws.append(RESULTS_HEADER)
    _save(wb, path)


def find_first_ticket(path: str, group: str, surname: str, name: str) -> int | None:
    """Номер билета из самой первой записи по студенту или None."""
    if not os.path.isfile(path):
        return None
    key = _key(group, surname, name)
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except PermissionError as exc:
        raise ResultsLockedError(str(exc)) from exc
    try:
        ws = wb.worksheets[0]
        for row in ws.iter_rows(min_row=2, max_col=4, values_only=True):
            if len(row) < 4:
                continue
            if _key(row[0], row[1], row[2]) == key and row[3] is not None:
                try:
                    return int(row[3])
                except (TypeError, ValueError):
                    continue
    finally:
        wb.close()
    return None


def append_result(path: str, group: str, surname: str, name: str,
                  ticket: int, is_repeat: bool, when: datetime | None = None) -> None:
    """Дописывает одну строку в конец журнала. Старые строки не трогаются."""
    ensure_results(path)
    try:
        wb = load_workbook(path)
    except PermissionError as exc:
        raise ResultsLockedError(str(exc)) from exc
    ws = wb.worksheets[0]
    when = when or datetime.now()
    ws.append([group, surname, name, ticket, when.strftime(DATE_FMT), "да" if is_repeat else "нет"])
    _save(wb, path)


def _save(wb, path: str) -> None:
    try:
        wb.save(path)
    except PermissionError as exc:
        raise ResultsLockedError(
            f"Файл {os.path.basename(path)} занят (возможно, открыт в Excel)") from exc


def _key(group, surname, name) -> tuple[str, str, str]:
    return tuple(_cell_str(v).casefold() for v in (group, surname, name))


# ---------------------------------------------------------------- бизнес-логика

def assign_ticket(results_path: str, tickets: dict[int, list[str]],
                  group: str, surname: str, name: str,
                  rng: random.Random | None = None,
                  when: datetime | None = None) -> Assignment:
    """Выдаёт билет и сразу пишет строку в журнал.

    Студент уже в журнале → тот же номер, «Повтор» = да.
    Иначе → случайный билет из распознанных, «Повтор» = нет.
    """
    first = find_first_ticket(results_path, group, surname, name)
    if first is not None:
        result = Assignment(first, True)
    else:
        if not tickets:
            raise ValueError("Нет ни одного корректного билета")
        rng = rng or random.Random()
        result = Assignment(rng.choice(sorted(tickets)), False)
    append_result(results_path, group, surname, name,
                  result.ticket_number, result.is_repeat, when)
    return result
