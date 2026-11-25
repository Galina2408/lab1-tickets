"""Лабораторная работа №1 — консольный генератор экзаменационных билетов."""
import os
import sys

from console_input import EscPressed, ask_non_empty, read_line
from journal import JournalLockedError, append_record, generate_ticket

JOURNAL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "journal.xlsx")


def save_with_retry(last_name: str, first_name: str, ticket: int) -> bool:
    """Сохраняет запись; если файл занят — просит закрыть его и пробует снова."""
    while True:
        try:
            append_record(JOURNAL_PATH, last_name, first_name, ticket)
            return True
        except JournalLockedError as exc:
            print(f"\n[!] {exc}")
            try:
                read_line("Закройте journal.xlsx в Excel и нажмите Enter для повтора "
                          "(ESC — выйти без сохранения этой записи): ")
            except EscPressed:
                return False


def main() -> None:
    if sys.platform != "win32":
        print("Приложение рассчитано на Windows (нужен модуль msvcrt).")
        return
    print("Для выхода нажмите ESC.\n")
    while True:
        try:
            last_name = ask_non_empty("Last name: ")
            first_name = ask_non_empty("First name: ")
        except EscPressed:
            print("Выход.")
            return
        ticket = generate_ticket()
        print(f"Билет № {ticket}")
        try:
            if save_with_retry(last_name, first_name, ticket):
                print("Записано в journal.xlsx\n")
            else:
                print("Запись не сохранена. Выход.")
                return
        except Exception as exc:  # любая другая ошибка файла — не падаем
            print(f"[!] Не удалось записать файл: {exc}\n")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nВыход.")
