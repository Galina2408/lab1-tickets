"""Посимвольный ввод строки с поддержкой ESC, Enter и Backspace.

На Windows клавиши читаются через msvcrt.getwch().
"""
from __future__ import annotations

import sys
from typing import Callable

ESC = "\x1b"
ENTER = ("\r", "\n")
BACKSPACE = ("\b", "\x7f")


class EscPressed(Exception):
    """Пользователь нажал ESC — нужно завершить программу."""


def _default_getch() -> str:
    import msvcrt  # только Windows
    return msvcrt.getwch()


def read_line(prompt: str,
              getch: Callable[[], str] | None = None,
              write: Callable[[str], object] | None = None) -> str:
    """Печатает prompt и собирает строку до Enter. ESC → EscPressed."""
    getch = getch or _default_getch
    write = write or _write
    write(prompt)
    buf: list[str] = []
    while True:
        ch = getch()
        if ch == ESC:
            write("\n")
            raise EscPressed
        if ch in ENTER:
            write("\n")
            return "".join(buf)
        if ch in BACKSPACE:
            if buf:
                buf.pop()
                write("\b \b")
            continue
        if ch in ("\x00", "\xe0"):  # стрелки, F1..F12 — второй код игнорируем
            getch()
            continue
        if ch == "\x03":  # Ctrl+C
            raise KeyboardInterrupt
        if ch.isprintable():
            buf.append(ch)
            write(ch)


def ask_non_empty(prompt: str, **kwargs) -> str:
    """Повторяет запрос, пока не введут непустую строку. Пробелы по краям обрезаются."""
    while True:
        value = read_line(prompt, **kwargs).strip()
        if value:
            return value
        (kwargs.get("write") or _write)("Поле не может быть пустым.\n")


def _write(text: str) -> None:
    sys.stdout.write(text)
    sys.stdout.flush()
