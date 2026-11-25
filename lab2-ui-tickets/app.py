"""Лабораторная работа №2 — генератор экзаменационных билетов (Tkinter)."""
from __future__ import annotations

import logging
import os
import sys
import tkinter as tk
from tkinter import messagebox, ttk

from ticket_core.data import (
    ResultsLockedError, SourceFileError, assign_ticket, ensure_results,
    load_students, load_tickets, students_of_group,
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
STUDENTS_PATH = os.path.join(DATA_DIR, "students.xlsx")
TICKETS_PATH = os.path.join(DATA_DIR, "tickets.docx")
RESULTS_PATH = os.path.join(DATA_DIR, "results.xlsx")
LOG_PATH = os.path.join(BASE_DIR, "app.log")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[logging.FileHandler(LOG_PATH, encoding="utf-8"), logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("tickets")


class TicketApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Генератор билетов")
        self.geometry("460x260")
        self.resizable(False, False)

        self.groups: dict[str, list[tuple[str, str]]] = {}
        self.tickets: dict[int, list[str]] = {}
        self.group_var = tk.StringVar()
        self.student_var = tk.StringVar()
        self.status_var = tk.StringVar()

        frm = ttk.Frame(self, padding=16)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(1, weight=1)

        ttk.Label(frm, text="Группа:").grid(row=0, column=0, sticky="w", pady=6)
        self.group_cb = ttk.Combobox(frm, textvariable=self.group_var, state="readonly")
        self.group_cb.grid(row=0, column=1, sticky="ew", pady=6)
        self.group_cb.bind("<<ComboboxSelected>>", self.on_group_selected)

        ttk.Label(frm, text="Студент:").grid(row=1, column=0, sticky="w", pady=6)
        self.student_cb = ttk.Combobox(frm, textvariable=self.student_var, state="readonly")
        self.student_cb.grid(row=1, column=1, sticky="ew", pady=6)
        self.student_cb.bind("<<ComboboxSelected>>", lambda _e: self.update_button())

        self.gen_btn = ttk.Button(frm, text="Сгенерировать билет",
                                  command=self.on_generate, state="disabled")
        self.gen_btn.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(14, 6))

        ttk.Button(frm, text="Перечитать файлы", command=self.load_data).grid(
            row=3, column=0, columnspan=2, sticky="ew")
        ttk.Label(frm, textvariable=self.status_var, foreground="gray",
                  wraplength=420).grid(row=4, column=0, columnspan=2, sticky="w", pady=(10, 0))

        self.after(50, self.load_data)

    # ---------------------------------------------------------------- данные
    def load_data(self):
        errors = []
        try:
            self.groups = load_students(STUDENTS_PATH)
        except SourceFileError as exc:
            self.groups = {}
            errors.append(str(exc))
        try:
            self.tickets = load_tickets(TICKETS_PATH)
            if not self.tickets:
                errors.append("В tickets.docx не найдено ни одного корректного билета.")
        except SourceFileError as exc:
            self.tickets = {}
            errors.append(str(exc))
        try:
            ensure_results(RESULTS_PATH)
        except ResultsLockedError as exc:
            errors.append(str(exc))

        self.group_cb["values"] = list(self.groups)
        self.group_var.set("")
        self.student_cb["values"] = []
        self.student_var.set("")
        self.update_button()
        self.status_var.set(f"Групп: {len(self.groups)}, билетов: {len(self.tickets)}")

        for e in errors:
            log.error(e)
        if errors:
            messagebox.showerror("Ошибка входных данных", "\n\n".join(errors) +
                                 "\n\nПоложите файлы в папку data и нажмите «Перечитать файлы».")

    # ---------------------------------------------------------------- UI-логика
    def on_group_selected(self, _event=None):
        students = students_of_group(self.groups, self.group_var.get())
        self.student_cb["values"] = students
        self.student_var.set("")
              self.status_var.set(f"Студентов в группе: {len(students)}" if students
                            else "В этой группе нет студентов.")
        self.update_button()

    def update_button(self):
        ready = bool(self.group_var.get()) and bool(self.student_var.get()) and bool(self.tickets)
        self.gen_btn.config(state="normal" if ready else "disabled")

    def _selected_student(self) -> tuple[str, str, str] | None:
        group = self.group_var.get()
        idx = self.student_cb.current()
        students = self.groups.get(group, [])
        if not group or idx < 0 or idx >= len(students):
            return None
        surname, name = students[idx]
        return group, surname, name

    def on_generate(self):
        sel = self._selected_student()
        if sel is None:
            return
        group, surname, name = sel
        while True:
            try:
                result = assign_ticket(RESULTS_PATH, self.tickets, group, surname, name)
                break
            except ResultsLockedError as exc:
                log.warning("results.xlsx заблокирован: %s", exc)
                if not messagebox.askretrycancel(
                        "Файл занят",
                        "Не удаётся записать results.xlsx — вероятно, он открыт в Excel.\n"
                        "Закройте файл и нажмите «Повторить».", parent=self):
                    return  # билет не выдан, журнал не изменён
            except Exception as exc:  # не роняем приложение
                log.exception("Ошибка генерации")
                messagebox.showerror("Ошибка", str(exc), parent=self)
                return
        log.info("%s %s %s → билет %d (повтор: %s)", group, surname, name,
                 result.ticket_number, result.is_repeat)
        self.show_result(f"{surname} {name}", result.ticket_number, result.is_repeat)

    def show_result(self, student: str, number: int, repeat: bool):
        win = tk.Toplevel(self)
        win.title("Билет")
        win.transient(self)
        win.resizable(False, False)
        body = ttk.Frame(win, padding=20)
        body.pack(fill="both", expand=True)

        ttk.Label(body, text=student).pack(anchor="w")
        ttk.Label(body, text=f"Ваш билет № {number}",
                  font=("Segoe UI", 16, "bold")).pack(anchor="w", pady=(4, 12))
        questions = self.tickets.get(number)
        if questions:
            for i, q in enumerate(questions, 1):
                ttk.Label(body, text=f"{i}. {q}", wraplength=480,
                          justify="left").pack(anchor="w", pady=3)
        else:
            ttk.Label(body, text="Вопросы этого билета не найдены в текущем tickets.docx.",
                      foreground="red").pack(anchor="w")
        if repeat:
            ttk.Label(body, text="Повторный запрос — выдан ранее назначенный билет.",
                      foreground="gray").pack(anchor="w", pady=(10, 0))
        ttk.Label(body, text="Esc — закрыть", foreground="gray").pack(anchor="e", pady=(14, 0))

        win.bind("<Escape>", lambda _e: win.destroy())
        win.grab_set()
        win.focus_force()


if __name__ == "__main__":
    TicketApp().mainloop()
