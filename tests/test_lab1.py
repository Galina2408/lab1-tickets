import os
import random
import sys
from datetime import datetime
from unittest import mock

import pytest
from openpyxl import load_workbook

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from console_input import EscPressed, ask_non_empty, read_line  # noqa: E402
from journal import HEADER, JournalLockedError, append_record, generate_ticket  # noqa: E402


def keys(text):
    it = iter(text)
    return lambda: next(it)


def silent(_):
    pass


def rows(path):
    return [list(r) for r in load_workbook(path).worksheets[0].iter_rows(values_only=True)]


# ---------------------------------------------------------------- ввод

def test_read_line_enter():
    assert read_line("> ", getch=keys("Ivanov\r"), write=silent) == "Ivanov"


def test_read_line_backspace():
    assert read_line("> ", getch=keys("Ivx\bov\r"), write=silent) == "Ivov"


def test_read_line_esc():
    with pytest.raises(EscPressed):
        read_line("> ", getch=keys("Iv\x1b"), write=silent)


def test_empty_input_is_repeated_and_trimmed():
    # пустая строка, строка из пробелов, затем нормальная
    assert ask_non_empty("> ", getch=keys("\r   \r  Petrov \r"), write=silent) == "Petrov"


# ---------------------------------------------------------------- билет

def test_ticket_in_range():
    rng = random.Random(0)
    assert all(1 <= generate_ticket(rng) <= 20 for _ in range(500))


# ---------------------------------------------------------------- журнал

def test_creates_file_with_header(tmp_path):
    p = str(tmp_path / "journal.xlsx")
    append_record(p, "Ivanov", "Ivan", 7, datetime(2026, 11, 25, 12, 0, 0))
    assert rows(p) == [HEADER, ["Ivanov", "Ivan", 7, "2026-11-25 12:00:00"]]


def test_appends_without_changing_old_rows(tmp_path):
    p = str(tmp_path / "journal.xlsx")
    for i in range(3):
        append_record(p, f"L{i}", f"F{i}", i + 1)
    before = rows(p)
    for i in range(3, 5):
        append_record(p, f"L{i}", f"F{i}", i + 1)
    after = rows(p)
    assert len(after) == 6  # шапка + 5
    assert after[:4] == before


def test_locked_file_raises_friendly_error(tmp_path):
    p = str(tmp_path / "journal.xlsx")
    append_record(p, "A", "B", 1)
    with mock.patch("openpyxl.workbook.workbook.Workbook.save",
                    side_effect=PermissionError("locked")):
        with pytest.raises(JournalLockedError):
            append_record(p, "C", "D", 2)
    assert len(rows(p)) == 2
