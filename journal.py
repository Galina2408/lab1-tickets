"""Работа с журналом journal.xlsx: создание, дописывание строк."""
from __future__ import annotations

import os
import random
from datetime import datetime

from openpyxl import Workbook, load_workbook

HEADER = ["Last name", "First name", "Номер билета", "Дата и время"]
DATE_FMT = "%Y-%m-%d %H:%M:%S"
MIN_TICKET, MAX_TICKET = 1, 20


class JournalLockedError(Exception):
    """journal.xlsx занят другой программой (открыт в Excel)."""


def generate_ticket(rng: random.Random | None = None) -> int:
    return (rng or random).randint(MIN_TICKET, MAX_TICKET)


def append_record(path: str, last_name: str, first_name: str, ticket: int,
                  when: datetime | None = None) -> None:
    """Дописывает строку в конец файла и сразу сохраняет.

    Если файла нет — создаёт его с шапкой. Старые строки не изменяются.
    """
    try:
        if os.path.isfile(path):
            wb = load_workbook(path)
            ws = wb.worksheets[0]
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = "Journal"
            ws.append(HEADER)
        when = when or datetime.now()
        ws.append([last_name, first_name, ticket, when.strftime(DATE_FMT)])
        wb.save(path)
    except PermissionError as exc:
        raise JournalLockedError(
            f"Файл {os.path.basename(path)} недоступен для записи — "
            f"вероятно, он открыт в Excel.") from exc
