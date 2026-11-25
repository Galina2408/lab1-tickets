import logging
import os
import random
import sys
from datetime import datetime
from unittest import mock

import pytest
from docx import Document
from openpyxl import Workbook, load_workbook

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from ticket_core import data  # noqa: E402
from ticket_core.data import (  # noqa: E402
    RESULTS_HEADER, ResultsLockedError, SourceFileError, assign_ticket,
    ensure_results, load_students, load_tickets, students_of_group,
)


# ---------------------------------------------------------------- фикстуры

@pytest.fixture
def students_xlsx(tmp_path):
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet("ИТ-21")
    ws.append(["Фамилия", "Имя"])
    ws.append(["Иванов", "Иван"])
    ws.append(["Петрова", "Анна"])
    ws.append([None, None])  # пустая строка — игнорируется
    ws2 = wb.create_sheet("ИТ-22")
    ws2.append(["Фамилия", "Имя"])
    ws2.append(["Попов", "Дмитрий"])
    empty = wb.create_sheet("Пустая")
    empty.append(["Фамилия", "Имя"])
    path = tmp_path / "students.xlsx"
    wb.save(path)
    return str(path)


def make_docx(tmp_path, lines):
    doc = Document()
    for line in lines:
        doc.add_paragraph(line)
    path = tmp_path / "tickets.docx"
    doc.save(path)
    return str(path)


@pytest.fixture
def tickets_docx(tmp_path):
    return make_docx(tmp_path, [
        "Билет 1", "1. Вопрос A", "2. Вопрос B", "3. Вопрос C", "",
        "Билет 2", "1. Вопрос D", "2. Вопрос E", "3. Вопрос F",
    ])


@pytest.fixture
def results_path(tmp_path):
    return str(tmp_path / "results.xlsx")


def rows(path):
    wb = load_workbook(path)
    return [list(r) for r in wb.worksheets[0].iter_rows(values_only=True)]


# ---------------------------------------------------------------- парсинг Word

def test_parse_tickets(tickets_docx):
    t = load_tickets(tickets_docx)
    assert t == {1: ["Вопрос A", "Вопрос B", "Вопрос C"],
                 2: ["Вопрос D", "Вопрос E", "Вопрос F"]}


def test_ticket_with_less_than_three_questions_is_skipped(tmp_path, caplog):
    path = make_docx(tmp_path, [
        "Билет 1", "1. A", "2. B",
        "Билет 2", "1. C", "2. D", "3. E",
    ])
    with caplog.at_level(logging.ERROR, logger="tickets"):
        t = load_tickets(path)
    assert list(t) == [2]
    assert "Билет 1 пропущен" in caplog.text


def test_unrecognized_header_is_skipped(tmp_path, caplog):
    path = make_docx(tmp_path, [
        "Билет первый", "1. A", "2. B", "3. C",
        "Билет 2", "1. C", "2. D", "3. E",
    ])
    with caplog.at_level(logging.ERROR, logger="tickets"):
        t = load_tickets(path)
    assert list(t) == [2]
    assert "Нераспознанный заголовок" in caplog.text


def test_word_auto_numbering_is_recognized():
    t = data.parse_tickets_paragraphs([
        ("Билет 3", False), ("Первый", True), ("Второй", True), ("Третий", True),
    ])
    assert t == {3: ["Первый", "Второй", "Третий"]}


def test_missing_tickets_file(tmp_path):
    with pytest.raises(SourceFileError):
        load_tickets(str(tmp_path / "nope.docx"))


# ---------------------------------------------------------------- студенты

def test_students_by_group(students_xlsx):
    groups = load_students(students_xlsx)
    assert list(groups) == ["ИТ-21", "ИТ-22", "Пустая"]
    assert students_of_group(groups, "ИТ-21") == ["Иванов Иван", "Петрова Анна"]
    assert students_of_group(groups, "ИТ-22") == ["Попов Дмитрий"]


def test_empty_group_returns_empty_list(students_xlsx):
    groups = load_students(students_xlsx)
    assert students_of_group(groups, "Пустая") == []
    assert students_of_group(groups, "Несуществующая") == []


def test_missing_students_file(tmp_path):
    with pytest.raises(SourceFileError):
        load_students(str(tmp_path / "nope.xlsx"))


# ---------------------------------------------------------------- журнал / идемпотентность

def test_results_created_with_header(results_path):
    ensure_results(results_path)
    assert rows(results_path) == [RESULTS_HEADER]


def test_repeat_returns_same_ticket_and_appends_row(results_path, tickets_docx):
    tickets = load_tickets(tickets_docx)
    t1 = datetime(2026, 11, 25, 10, 0, 0)
    t2 = datetime(2026, 11, 25, 10, 5, 0)

    first = assign_ticket(results_path, tickets, "ИТ-21", "Иванов", "Иван",
                          rng=random.Random(1), when=t1)
    assert first.is_repeat is False
    snapshot = rows(results_path)

    # другой rng — номер всё равно должен совпасть
    second = assign_ticket(results_path, tickets, "ИТ-21", "Иванов", "Иван",
                           rng=random.Random(999), when=t2)
    assert second.ticket_number == first.ticket_number
    assert second.is_repeat is True

    after = rows(results_path)
    assert after[:len(snapshot)] == snapshot          # старые строки не тронуты
    assert len(after) == 3
    assert after[1] == ["ИТ-21", "Иванов", "Иван", first.ticket_number, "2026-11-25 10:00:00", "нет"]
    assert after[2] == ["ИТ-21", "Иванов", "Иван", first.ticket_number, "2026-11-25 10:05:00", "да"]


def test_same_name_other_group_is_new_student(results_path, tickets_docx):
    tickets = load_tickets(tickets_docx)
    assign_ticket(results_path, tickets, "ИТ-21", "Иванов", "Иван")
    other = assign_ticket(results_path, tickets, "ИТ-22", "Иванов", "Иван")
    assert other.is_repeat is False


def test_random_ticket_in_range(results_path, tickets_docx):
    tickets = load_tickets(tickets_docx)
    for i in range(20):
        r = assign_ticket(results_path, tickets, "Г", f"Студент{i}", "Имя")
        assert r.ticket_number in tickets


def test_locked_results_raises_friendly_error(results_path, tickets_docx):
    tickets = load_tickets(tickets_docx)
    ensure_results(results_path)
    with mock.patch("openpyxl.workbook.workbook.Workbook.save",
                    side_effect=PermissionError("locked")):
        with pytest.raises(ResultsLockedError):
            assign_ticket(results_path, tickets, "ИТ-21", "Иванов", "Иван")
    assert rows(results_path) == [RESULTS_HEADER]  # журнал не испорчен
