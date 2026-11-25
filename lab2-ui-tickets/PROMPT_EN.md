# Step 1 — Prompt for Copilot Chat (English spec)

Paste this into Copilot Chat (VS Code) to generate / refine the project.

---

**Role:** You are a senior Python developer. Build a desktop GUI app with **Python 3.10+ and Tkinter (ttk)**. Dependencies: `openpyxl`, `python-docx`, `pytest`. No other libraries.

**Project layout**
```
app.py                 # Tkinter GUI only
ticket_core/data.py    # all file I/O and business logic, no GUI imports
tests/test_core.py     # pytest unit tests
data/students.xlsx, data/tickets.docx, data/results.xlsx
```

**Inputs**
- `students.xlsx`: one sheet per group, sheet name = group name. Row 1 is a header. From row 2: column A = surname, column B = first name. Skip fully empty rows.
- `tickets.docx`: a header paragraph `Билет N` followed by numbered question paragraphs `N. text`. Also accept Word auto-numbered list paragraphs (numPr) as questions.

**Parsing rules**
- A ticket with fewer than 3 questions, a duplicate number, or an unrecognized header (e.g. `Билет первый`) is skipped. Log the error via `logging` and continue with the rest. If more than 3 questions, keep the first 3 and log a warning.

**GUI behaviour**
1. On start, combobox #1 (readonly) is filled with the sheet names (groups).
2. After a group is selected, combobox #2 is filled with `"Surname Name"` of that group, and its selection is cleared.
3. The "Сгенерировать билет" button is disabled until both a group and a student are selected. An empty group → empty list, button stays disabled, no exception.
4. Clicking the button opens a modal Toplevel showing `Ваш билет № N` and the 3 questions.
5. `Esc` closes only the result window; the main window stays open for the next student.

**Ticket assignment (idempotent)**
- Key = (group, surname, name), compared trimmed and case-insensitive.
- If the key already exists in `results.xlsx`, reuse the ticket number from the **first** matching row; append a new row with `Повтор = "да"`.
- Otherwise pick a random ticket among parsed tickets (1..N) and append a row with `Повтор = "нет"`.
- Make the RNG and the timestamp injectable for tests.

**results.xlsx**
- Created with header `Группа | Фамилия | Имя | Номер билета | Дата и время | Повтор (да/нет)` if missing.
- Append-only: never modify or rewrite existing rows. Date format `YYYY-MM-DD HH:MM:SS`.
- Save immediately after each generation (not on app exit).

**Error handling (app must never crash)**
- Missing/corrupt `students.xlsx` or `tickets.docx` → `messagebox.showerror` with a clear message; app stays open; provide a "Перечитать файлы" button.
- `PermissionError` while saving `results.xlsx` (file open in Excel) → raise a custom `ResultsLockedError`; the GUI shows `askretrycancel` and retries in a loop until success or cancel.

**Tests (pytest, use `tmp_path`, generate files in tests)**
- Parse a valid ticket from Word; skip ticket with < 3 questions; skip bad header.
- Students by group; empty group returns `[]`.
- Idempotency: the same student twice → same ticket number, 2 rows, second has `"да"`, first row unchanged.
- Locked file → `ResultsLockedError`, journal untouched (mock `Workbook.save`).

Output complete files, no placeholders.
