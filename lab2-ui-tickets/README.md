# Лабораторная работа №2 — UI-генератор билетов

Десктопное приложение на **Python + Tkinter**. Вход: `data/students.xlsx`, `data/tickets.docx`. Выход: `data/results.xlsx` (журнал, только дописывание).

## Структура

```
app.py                 — интерфейс (Tkinter)
ticket_core/data.py    — чтение Excel/Word, журнал, выдача билета (без GUI)
tests/test_core.py     — unit-тесты (pytest), 13 штук
make_sample_data.py    — создаёт примеры students.xlsx и tickets.docx
data/                  — входные файлы и results.xlsx
PROMPT_EN.md           — ТЗ на английском для Copilot Chat (Шаг 1)
```

## Установка (Windows)

Git, VS Code и Copilot уже настроены после Лаб 1. Нужен только Python.

1. **Python 3.10+**: скачать с https://www.python.org/downloads/ и при установке **поставить галочку «Add python.exe to PATH»**. Tkinter входит в стандартный установщик.
   Проверка в Git Bash: `python --version`
2. **Расширение Python для VS Code**: Extensions (Ctrl+Shift+X) → «Python» (Microsoft) → Install.
3. Распаковать папку проекта **внутрь своего репозитория из Лаб 1** и открыть её в VS Code (File → Open Folder).
4. В терминале VS Code (Ctrl+`; выбрать Git Bash):

```bash
python -m venv .venv
source .venv/Scripts/activate      # в PowerShell: .venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

5. VS Code: Ctrl+Shift+P → «Python: Select Interpreter» → выбрать `.venv`.

## Запуск

```bash
python make_sample_data.py   # только если нужны тестовые файлы (перезапишет data/students.xlsx и tickets.docx)
python app.py
```

## Тесты

```bash
python -m pytest -v
```

Покрыто: парсинг билета из Word, пропуск билетов с <3 вопросами и с кривым заголовком, автонумерация Word, студенты по группе, пустая группа, отсутствующие файлы, идемпотентность (повтор → тот же билет, новая строка «да», старые строки не тронуты), заблокированный results.xlsx.

## Как проверить критерии приёмки вручную

| Критерий | Что сделать |
|---|---|
| Список студентов = лист | Выбрать «ИТ-21», затем «ИТ-22» |
| Кнопка только при обоих выборах | До выбора студента кнопка серая; группа «ИТ-23 (пустая)» → список пуст, кнопка серая |
| Билет и 3 вопроса | Нажать кнопку |
| ESC | Esc закрывает окно билета, главное окно остаётся |
| Повтор | Выбрать того же студента ещё раз → тот же номер, в журнале новая строка «да» |
| Запись сразу | Открыть `data/results.xlsx` после генерации |
| Блокировка | Открыть results.xlsx в Excel, нажать кнопку → «Файл занят», закрыть Excel → «Повторить» |
| Нет файла | Переименовать tickets.docx, запустить → сообщение, приложение работает |
| Кривые билеты | В примере есть «Билет 6» (1 вопрос) и «Билет без номера» — пропускаются, см. `app.log` |

## Шаг 4 — Git

```bash
git checkout -b lab2-ui-tickets
git add .
git commit -m "Lab 2: UI-генератор билетов"
git push -u origin lab2-ui-tickets
```

`-u origin lab2-ui-tickets` нужен при первом push новой ветки (без него Git попросит указать upstream).

## Примечания

- Студент определяется тройкой «группа + фамилия + имя» (без учёта регистра и пробелов по краям). Однофамилец в другой группе — другой студент.
- Случайный билет выбирается только из корректно распознанных билетов.
- Если запись в журнал не удалась и нажата «Отмена» — билет не показывается, журнал не меняется.
