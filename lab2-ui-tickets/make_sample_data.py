"""Создаёт примеры data/students.xlsx и data/tickets.docx для проверки."""
import os

from docx import Document
from openpyxl import Workbook

DATA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
os.makedirs(DATA, exist_ok=True)

wb = Workbook()
wb.remove(wb.active)
groups = {
    "ИТ-21": [("Иванов", "Иван"), ("Петрова", "Анна"), ("Сидоров", "Павел")],
    "ИТ-22": [("Кузнецова", "Мария"), ("Попов", "Дмитрий")],
    "ИТ-23 (пустая)": [],
}
for title, students in groups.items():
    ws = wb.create_sheet(title)
    ws.append(["Фамилия", "Имя"])
    for s in students:
        ws.append(list(s))
wb.save(os.path.join(DATA, "students.xlsx"))

doc = Document()
topics = [
    ("Что такое Git?", "Ветки и слияние", "Команда git push"),
    ("Типы данных Python", "Списки и кортежи", "Словари"),
    ("ООП: классы", "Наследование", "Полиморфизм"),
    ("Исключения", "Контекстные менеджеры", "Логирование"),
    ("Модульное тестирование", "pytest: фикстуры", "Моки"),
]
for i, qs in enumerate(topics, 1):
    doc.add_paragraph(f"Билет {i}")
    for j, q in enumerate(qs, 1):
        doc.add_paragraph(f"{j}. {q}")
    doc.add_paragraph("")
# Намеренно некорректные билеты — должны пропускаться с записью в лог
doc.add_paragraph("Билет 6")
doc.add_paragraph("1. Только один вопрос")
doc.add_paragraph("Билет без номера")
doc.add_paragraph("1. Вопрос")
doc.add_paragraph("2. Вопрос")
doc.add_paragraph("3. Вопрос")
doc.save(os.path.join(DATA, "tickets.docx"))
print("Готово: data/students.xlsx, data/tickets.docx")
