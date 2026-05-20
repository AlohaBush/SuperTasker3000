# Public GitHub version. Do not commit telegram_config.json with real tokens.
# Configure secrets via telegram_config.json locally or environment variables.
import argparse
import html
import json
import re
import shutil
import subprocess
import os
import sys
import tempfile
import time
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from difflib import SequenceMatcher
from pathlib import Path
from uuid import uuid4
from zipfile import ZIP_DEFLATED, ZipFile


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_CONFIG = BASE_DIR / "telegram_config.json"
DEFAULT_PLANNER = BASE_DIR / "Task_Planner.xlsx"
STATE_FILE = BASE_DIR / "telegram_task_bot_state.json"
SUPPORTED_TRANSCRIPTION_SUFFIXES = {".mp3", ".mp4", ".mpeg", ".mpga", ".m4a", ".wav", ".webm"}
RU_MONTHS = {
    "января": 1,
    "январь": 1,
    "февраля": 2,
    "февраль": 2,
    "марта": 3,
    "март": 3,
    "апреля": 4,
    "апрель": 4,
    "мая": 5,
    "май": 5,
    "июня": 6,
    "июнь": 6,
    "июля": 7,
    "июль": 7,
    "августа": 8,
    "август": 8,
    "сентября": 9,
    "сентябрь": 9,
    "октября": 10,
    "октябрь": 10,
    "ноября": 11,
    "ноябрь": 11,
    "декабря": 12,
    "декабрь": 12,
}
RU_WEEKDAYS = {
    "понедельник": 0,
    "понедельника": 0,
    "вторник": 1,
    "вторника": 1,
    "среду": 2,
    "среда": 2,
    "среды": 2,
    "четверг": 3,
    "четверга": 3,
    "пятницу": 4,
    "пятница": 4,
    "пятницы": 4,
    "субботу": 5,
    "суббота": 5,
    "субботы": 5,
    "воскресенье": 6,
    "воскресенья": 6,
}
NEXT_WEEK_CONTEXT_RE = r"(?:следующ\w*|будущ\w*|след\.?)\s+недел\w*"
CURRENT_WEEK_CONTEXT_RE = r"(?:эт\w*|текущ\w*)\s+недел\w*"
WEEK_CONTEXT_RE = rf"(?:{NEXT_WEEK_CONTEXT_RE}|{CURRENT_WEEK_CONTEXT_RE})"
NEXT_WEEKDAY_PREFIX_RE = r"(?:следующ\w*|будущ\w*|след\.?)"
RU_DAY_WORDS = {
    "первое": 1, "первого": 1,
    "второе": 2, "второго": 2,
    "третье": 3, "третьего": 3,
    "четвертое": 4, "четвертого": 4,
    "пятое": 5, "пятого": 5,
    "шестое": 6, "шестого": 6,
    "седьмое": 7, "седьмого": 7,
    "восьмое": 8, "восьмого": 8,
    "девятое": 9, "девятого": 9,
    "десятое": 10, "десятого": 10,
    "одиннадцатое": 11, "одиннадцатого": 11,
    "двенадцатое": 12, "двенадцатого": 12,
    "тринадцатое": 13, "тринадцатого": 13,
    "четырнадцатое": 14, "четырнадцатого": 14,
    "пятнадцатое": 15, "пятнадцатого": 15,
    "шестнадцатое": 16, "шестнадцатого": 16,
    "семнадцатое": 17, "семнадцатого": 17,
    "восемнадцатое": 18, "восемнадцатого": 18,
    "девятнадцатое": 19, "девятнадцатого": 19,
    "двадцатое": 20, "двадцатого": 20,
    "двадцать первое": 21, "двадцать первого": 21,
    "двадцать второе": 22, "двадцать второго": 22,
    "двадцать третье": 23, "двадцать третьего": 23,
    "двадцать четвертое": 24, "двадцать четвертого": 24,
    "двадцать пятое": 25, "двадцать пятого": 25,
    "двадцать шестое": 26, "двадцать шестого": 26,
    "двадцать седьмое": 27, "двадцать седьмого": 27,
    "двадцать восьмое": 28, "двадцать восьмого": 28,
    "двадцать девятое": 29, "двадцать девятого": 29,
    "тридцатое": 30, "тридцатого": 30,
    "тридцать первое": 31, "тридцать первого": 31,
}
RU_NUMBER_WORDS = {
    "ноль": 0,
    "один": 1,
    "одна": 1,
    "первый": 1,
    "первую": 1,
    "первая": 1,
    "два": 2,
    "две": 2,
    "второй": 2,
    "вторую": 2,
    "вторая": 2,
    "три": 3,
    "третий": 3,
    "третью": 3,
    "третья": 3,
    "четыре": 4,
    "четвертый": 4,
    "четвертую": 4,
    "четвертая": 4,
    "пять": 5,
    "пятый": 5,
    "пятую": 5,
    "пятая": 5,
    "шесть": 6,
    "шестой": 6,
    "шестую": 6,
    "шестая": 6,
    "семь": 7,
    "седьмой": 7,
    "седьмую": 7,
    "седьмая": 7,
    "восемь": 8,
    "восьмой": 8,
    "восьмую": 8,
    "восьмая": 8,
    "девять": 9,
    "девятый": 9,
    "девятую": 9,
    "девятая": 9,
    "десять": 10,
    "одиннадцать": 11,
    "двенадцать": 12,
    "тринадцать": 13,
    "четырнадцать": 14,
    "пятнадцать": 15,
    "шестнадцать": 16,
    "семнадцать": 17,
    "восемнадцать": 18,
    "девятнадцать": 19,
    "двадцать": 20,
    "тридцать": 30,
    "сорок": 40,
    "пятьдесят": 50,
    "шестьдесят": 60,
    "семьдесят": 70,
    "восемьдесят": 80,
    "девяносто": 90,
    "сто": 100,
    "двести": 200,
    "триста": 300,
    "четыреста": 400,
    "пятьсот": 500,
    "шестьсот": 600,
    "семьсот": 700,
    "восемьсот": 800,
    "девятьсот": 900,
}

NS = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
ET.register_namespace("", NS["x"])


@dataclass
class Task:
    row: int
    name: str
    due: date | None
    completed: bool = False
    completed_on: date | None = None


@dataclass
class DateInterval:
    start: date
    end: date
    title: str


def column_name(cell_ref: str) -> str:
    return "".join(ch for ch in cell_ref if ch.isalpha())


def excel_serial(day: date) -> int:
    return (day - date(1899, 12, 30)).days


def from_excel_serial(value: str) -> date | None:
    try:
        return date(1899, 12, 30) + timedelta(days=int(float(value)))
    except (TypeError, ValueError):
        return None


def parse_date(text: str, today: date | None = None) -> date:
    today = today or date.today()
    raw = text.strip().lower()
    shortcuts = {
        "сегодня": today,
        "завтра": today + timedelta(days=1),
        "послезавтра": today + timedelta(days=2),
        "today": today,
        "tomorrow": today + timedelta(days=1),
    }
    if raw in shortcuts:
        return shortcuts[raw]

    for fmt in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            pass
    raise ValueError("Не понял дату. Пример: 25.05.2026, 2026-05-25, сегодня, завтра.")


def extract_query_date(text: str, today: date | None = None) -> date | None:
    today = today or date.today()
    cleaned = text.lower().replace("ё", "е")

    if "послезавтра" in cleaned:
        return today + timedelta(days=2)
    if "завтра" in cleaned:
        return today + timedelta(days=1)
    if "сегодня" in cleaned:
        return today

    next_week_context = re.search(rf"\b{NEXT_WEEK_CONTEXT_RE}\b", cleaned)
    current_week_context = re.search(rf"\b{CURRENT_WEEK_CONTEXT_RE}\b", cleaned)
    current_week = today - timedelta(days=today.weekday())

    for weekday_word, weekday in RU_WEEKDAYS.items():
        if re.search(rf"\b{weekday_word}\b", cleaned):
            if next_week_context:
                return current_week + timedelta(days=7 + weekday)
            if current_week_context:
                return current_week + timedelta(days=weekday)
            delta = (weekday - today.weekday()) % 7
            if re.search(rf"\b{NEXT_WEEKDAY_PREFIX_RE}\s+{weekday_word}\b", cleaned):
                delta = delta or 7
            return today + timedelta(days=delta)

    patterns = [
        r"\b(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})\b",
        r"\b(\d{4})-(\d{1,2})-(\d{1,2})\b",
    ]
    first = re.search(patterns[0], cleaned)
    if first:
        day, month, year = map(int, first.groups())
        if year < 100:
            year += 2000
        return date(year, month, day)

    second = re.search(patterns[1], cleaned)
    if second:
        year, month, day = map(int, second.groups())
        return date(year, month, day)

    month_names = "|".join(RU_MONTHS)
    match = re.search(rf"\b(\d{{1,2}})\s+({month_names})(?:\s+(\d{{4}}))?\b", cleaned)
    if match:
        day = int(match.group(1))
        month = RU_MONTHS[match.group(2)]
        year = int(match.group(3)) if match.group(3) else today.year
        return date(year, month, day)

    day_words = "|".join(sorted(map(re.escape, RU_DAY_WORDS), key=len, reverse=True))
    word_match = re.search(rf"\b({day_words})\s+({month_names})(?:\s+(\d{{4}}))?\b", cleaned)
    if word_match:
        day = RU_DAY_WORDS[word_match.group(1)]
        month = RU_MONTHS[word_match.group(2)]
        year = int(word_match.group(3)) if word_match.group(3) else today.year
        return date(year, month, day)
    return None


def format_date(day: date | None) -> str:
    return day.strftime("%d.%m.%Y") if day else "без срока"


def html_text(value: object) -> str:
    return html.escape(str(value), quote=False)


def bold(value: object) -> str:
    return f"<b>{html_text(value)}</b>"


def status_for(task: Task, today: date | None = None) -> str:
    today = today or date.today()
    if task.completed:
        return f"выполнено {format_date(task.completed_on)}" if task.completed_on else "выполнено"
    if task.due is None:
        return "без срока"
    delta = (task.due - today).days
    if delta < 0:
        return f"просрочена на {abs(delta)} дн."
    if delta == 0:
        return "сегодня"
    if delta == 1:
        return "завтра"
    return f"через {delta} дн."


def status_html(task: Task, today: date | None = None) -> str:
    today = today or date.today()
    if not task.completed and task.due is not None and task.due < today:
        days = abs((task.due - today).days)
        return f"🔴 {bold('просрочена')} на {days} дн."
    return html_text(status_for(task, today))


def cell_text(cell: ET.Element) -> str:
    if cell.get("t") == "inlineStr":
        text_node = cell.find("x:is/x:t", NS)
        return text_node.text if text_node is not None and text_node.text else ""
    value_node = cell.find("x:v", NS)
    return value_node.text if value_node is not None and value_node.text else ""


def row_cells(row: ET.Element) -> dict[str, ET.Element]:
    cells = {}
    for cell in row.findall("x:c", NS):
        ref = cell.get("r", "")
        if ref:
            cells[column_name(ref)] = cell
    return cells


def load_task_sheet(planner_file: Path) -> tuple[ET.ElementTree, ET.Element, bytes, dict[str, bytes]]:
    with ZipFile(planner_file, "r") as workbook:
        sheet_bytes = workbook.read("xl/worksheets/sheet1.xml")
        other_files = {name: workbook.read(name) for name in workbook.namelist() if name != "xl/worksheets/sheet1.xml"}
    tree = ET.ElementTree(ET.fromstring(sheet_bytes))
    root = tree.getroot()
    sheet_data = root.find("x:sheetData", NS)
    if sheet_data is None:
        raise RuntimeError("Не нашел таблицу задач внутри Excel-файла.")
    return tree, sheet_data, sheet_bytes, other_files


def save_task_sheet(planner_file: Path, tree: ET.ElementTree, other_files: dict[str, bytes]) -> None:
    tmp_file = planner_file.with_suffix(".tmp.xlsx")
    xml_bytes = ET.tostring(tree.getroot(), encoding="utf-8", xml_declaration=True)
    with ZipFile(tmp_file, "w", ZIP_DEFLATED) as workbook:
        for name, data in other_files.items():
            workbook.writestr(name, data)
        workbook.writestr("xl/worksheets/sheet1.xml", xml_bytes)
    tmp_file.replace(planner_file)


def get_tasks(planner_file: Path = DEFAULT_PLANNER) -> list[Task]:
    _, sheet_data, _, _ = load_task_sheet(planner_file)
    tasks: list[Task] = []
    for row in sheet_data.findall("x:row", NS):
        row_number = int(row.get("r", "0"))
        if row_number < 2:
            continue
        cells = row_cells(row)
        name = cell_text(cells["A"]).strip() if "A" in cells else ""
        if not name:
            continue
        due = None
        if "B" in cells:
            raw_due = cell_text(cells["B"]).strip()
            due = from_excel_serial(raw_due)
            if due is None and raw_due:
                try:
                    due = parse_date(raw_due)
                except ValueError:
                    due = None
        completed = False
        if "F" in cells:
            completed_raw = cell_text(cells["F"]).strip().lower()
            completed = completed_raw in {"да", "yes", "true", "1", "done", "completed", "выполнено"}
        completed_on = None
        if "G" in cells:
            completed_on_raw = cell_text(cells["G"]).strip()
            completed_on = from_excel_serial(completed_on_raw)
        tasks.append(Task(row=row_number, name=name, due=due, completed=completed, completed_on=completed_on))
    return tasks


def find_or_create_row(sheet_data: ET.Element) -> ET.Element:
    rows = sheet_data.findall("x:row", NS)
    for row in rows:
        row_number = int(row.get("r", "0"))
        if row_number < 2:
            continue
        cells = row_cells(row)
        name = cell_text(cells["A"]).strip() if "A" in cells else ""
        if not name:
            return row

    max_row = max((int(row.get("r", "0")) for row in rows), default=1)
    new_number = max_row + 1
    new_row = ET.Element(f"{{{NS['x']}}}row", {"r": str(new_number), "ht": "21", "customHeight": "1"})
    sheet_data.append(new_row)
    return new_row


def make_cell(row_number: int, col: str) -> ET.Element:
    return ET.Element(f"{{{NS['x']}}}c", {"r": f"{col}{row_number}"})


def ensure_cell(row: ET.Element, col: str) -> ET.Element:
    row_number = int(row.get("r", "0"))
    cells = row_cells(row)
    if col in cells:
        return cells[col]
    cell = make_cell(row_number, col)
    row.append(cell)
    row[:] = sorted(row, key=lambda node: node.get("r", ""))
    return cell


def set_inline(cell: ET.Element, value: str, style: str = "0") -> None:
    ref = cell.get("r", "")
    cell.clear()
    cell.set("r", ref)
    cell.set("t", "inlineStr")
    cell.set("s", style)
    inline = ET.SubElement(cell, f"{{{NS['x']}}}is")
    text = ET.SubElement(inline, f"{{{NS['x']}}}t")
    text.text = value


def set_number(cell: ET.Element, value: int, style: str = "3") -> None:
    ref = cell.get("r", "")
    cell.clear()
    cell.set("r", ref)
    cell.set("s", style)
    node = ET.SubElement(cell, f"{{{NS['x']}}}v")
    node.text = str(value)


def set_formula(cell: ET.Element, formula: str, style: str = "4") -> None:
    ref = cell.get("r", "")
    cell.clear()
    cell.set("r", ref)
    cell.set("s", style)
    node = ET.SubElement(cell, f"{{{NS['x']}}}f")
    node.text = formula


def ensure_formula_cells(row: ET.Element) -> None:
    row_number = int(row.get("r", "0"))
    formulas = {
        "C": f'IF(A{row_number}="","",IF(F{row_number}="Да","Выполнена",IF(B{row_number}="","Нет срока",IF(INT(B{row_number})<TODAY(),"Просрочена",IF(INT(B{row_number})=TODAY(),"Сегодня",IF(INT(B{row_number})<=TODAY()+7,"Скоро","Запланирована"))))))',
        "D": f'IF(OR(A{row_number}="",B{row_number}="",F{row_number}="Да"),"",INT(B{row_number})-TODAY())',
        "E": f'IF(A{row_number}="","",IF(F{row_number}="Да","Выполнено",IF(B{row_number}="","Без срока",IF(INT(B{row_number})<TODAY(),"Просрочено",IF(INT(B{row_number})=TODAY(),"Сегодня",IF(INT(B{row_number})<=TODAY()+7,"Скоро","Позже"))))))',
    }
    for col, formula in formulas.items():
        cell = ensure_cell(row, col)
        set_formula(cell, formula)


def add_task(name: str, due: date, planner_file: Path = DEFAULT_PLANNER) -> Task:
    tree, sheet_data, _, other_files = load_task_sheet(planner_file)
    row = find_or_create_row(sheet_data)
    row_number = int(row.get("r", "0"))
    set_inline(ensure_cell(row, "A"), name)
    set_number(ensure_cell(row, "B"), excel_serial(due))
    clear_cell(ensure_cell(row, "F"), "4")
    clear_cell(ensure_cell(row, "G"), "3")
    ensure_formula_cells(row)
    save_task_sheet(planner_file, tree, other_files)
    return Task(row=row_number, name=name, due=due)


def clear_cell(cell: ET.Element, style: str | None = None) -> None:
    ref = cell.get("r", "")
    current_style = style if style is not None else cell.get("s")
    cell.clear()
    cell.set("r", ref)
    cell.set("t", "inlineStr")
    if current_style is not None:
        cell.set("s", current_style)
    inline = ET.SubElement(cell, f"{{{NS['x']}}}is")
    ET.SubElement(inline, f"{{{NS['x']}}}t")


def delete_task(row_number: int, planner_file: Path = DEFAULT_PLANNER) -> Task:
    if row_number < 2:
        raise ValueError("Укажите номер задачи из списка /list.")

    tree, sheet_data, _, other_files = load_task_sheet(planner_file)
    target_row = None
    for row in sheet_data.findall("x:row", NS):
        if int(row.get("r", "0")) == row_number:
            target_row = row
            break
    if target_row is None:
        raise ValueError(f"Не нашел задачу с номером {row_number}.")

    cells = row_cells(target_row)
    name = cell_text(cells["A"]).strip() if "A" in cells else ""
    if not name:
        raise ValueError(f"Строка {row_number} уже пустая.")

    due = None
    if "B" in cells:
        due = from_excel_serial(cell_text(cells["B"]).strip())

    clear_cell(ensure_cell(target_row, "A"), "0")
    clear_cell(ensure_cell(target_row, "B"), "3")
    clear_cell(ensure_cell(target_row, "F"), "4")
    clear_cell(ensure_cell(target_row, "G"), "3")
    ensure_formula_cells(target_row)
    save_task_sheet(planner_file, tree, other_files)
    return Task(row=row_number, name=name, due=due)


def get_task_by_row(row_number: int, planner_file: Path = DEFAULT_PLANNER) -> Task:
    for task in get_tasks(planner_file):
        if task.row == row_number:
            return task
    raise ValueError(f"Не нашел задачу с номером {row_number}.")


def move_tasks(row_numbers: list[int], new_due: date, planner_file: Path = DEFAULT_PLANNER) -> tuple[list[Task], list[int]]:
    tree, sheet_data, _, other_files = load_task_sheet(planner_file)
    rows_by_number = {int(row.get("r", "0")): row for row in sheet_data.findall("x:row", NS)}
    moved: list[Task] = []
    missing: list[int] = []

    for row_number in unique_ints(row_numbers):
        row = rows_by_number.get(row_number)
        if row is None:
            missing.append(row_number)
            continue
        cells = row_cells(row)
        name = cell_text(cells["A"]).strip() if "A" in cells else ""
        if not name:
            missing.append(row_number)
            continue
        set_number(ensure_cell(row, "B"), excel_serial(new_due))
        ensure_formula_cells(row)
        completed = cell_text(cells["F"]).strip().lower() in {"да", "yes", "true", "1", "done", "completed", "выполнено"} if "F" in cells else False
        moved.append(Task(row=row_number, name=name, due=new_due, completed=completed))

    if moved:
        save_task_sheet(planner_file, tree, other_files)
    return moved, missing


def rename_tasks(row_names: list[tuple[int, str]], planner_file: Path = DEFAULT_PLANNER) -> tuple[list[Task], list[int]]:
    tree, sheet_data, _, other_files = load_task_sheet(planner_file)
    rows_by_number = {int(row.get("r", "0")): row for row in sheet_data.findall("x:row", NS)}
    renamed: list[Task] = []
    missing: list[int] = []

    for row_number, new_name in row_names:
        row = rows_by_number.get(row_number)
        clean_name = new_name.strip()
        if row is None or not clean_name:
            missing.append(row_number)
            continue
        cells = row_cells(row)
        old_name = cell_text(cells["A"]).strip() if "A" in cells else ""
        if not old_name:
            missing.append(row_number)
            continue
        due = from_excel_serial(cell_text(cells["B"]).strip()) if "B" in cells else None
        completed = cell_text(cells["F"]).strip().lower() in {"да", "yes", "true", "1", "done", "completed", "выполнено"} if "F" in cells else False
        completed_on = from_excel_serial(cell_text(cells["G"]).strip()) if "G" in cells else None
        set_inline(ensure_cell(row, "A"), clean_name)
        ensure_formula_cells(row)
        renamed.append(Task(row=row_number, name=clean_name, due=due, completed=completed, completed_on=completed_on))

    if renamed:
        save_task_sheet(planner_file, tree, other_files)
    return renamed, missing


def complete_tasks(row_numbers: list[int], planner_file: Path = DEFAULT_PLANNER) -> tuple[list[Task], list[int]]:
    tree, sheet_data, _, other_files = load_task_sheet(planner_file)
    rows_by_number = {int(row.get("r", "0")): row for row in sheet_data.findall("x:row", NS)}
    completed_tasks: list[Task] = []
    missing: list[int] = []
    completed_on = date.today()

    for row_number in unique_ints(row_numbers):
        row = rows_by_number.get(row_number)
        if row is None:
            missing.append(row_number)
            continue
        cells = row_cells(row)
        name = cell_text(cells["A"]).strip() if "A" in cells else ""
        if not name:
            missing.append(row_number)
            continue
        due = from_excel_serial(cell_text(cells["B"]).strip()) if "B" in cells else None
        set_inline(ensure_cell(row, "F"), "Да", "4")
        set_number(ensure_cell(row, "G"), excel_serial(completed_on), "3")
        ensure_formula_cells(row)
        completed_tasks.append(Task(row=row_number, name=name, due=due, completed=True, completed_on=completed_on))

    if completed_tasks:
        save_task_sheet(planner_file, tree, other_files)
    return completed_tasks, missing


def select_tasks(kind: str, tasks: list[Task], soon_days: int) -> list[Task]:
    today = date.today()
    dated = [task for task in tasks if task.due is not None and not task.completed]
    if kind == "today":
        return sorted([task for task in dated if task.due == today], key=lambda task: task.due)
    if kind == "overdue":
        return sorted([task for task in dated if task.due < today], key=lambda task: task.due)
    if kind == "soon":
        limit = today + timedelta(days=soon_days)
        return sorted([task for task in dated if today < task.due <= limit], key=lambda task: task.due)
    if kind == "nearest":
        return sorted([task for task in dated if task.due >= today], key=lambda task: task.due)[:10]
    return sorted(dated, key=lambda task: task.due)


def select_tasks_for_date(tasks: list[Task], target_date: date) -> list[Task]:
    return sorted([task for task in tasks if task.due == target_date and not task.completed], key=lambda task: (task.due or date.max, task.row))


def select_tasks_for_interval(tasks: list[Task], interval: DateInterval) -> list[Task]:
    return sorted(
        [
            task
            for task in tasks
            if task.due is not None and interval.start <= task.due <= interval.end and not task.completed
        ],
        key=lambda task: (task.due or date.max, task.row),
    )


def render_tasks(title: str, tasks: list[Task]) -> str:
    if not tasks:
        return f"{html_text(title)}\nНет задач."
    task_lines = [
        f"- {html_text(task.name)} — {bold(format_date(task.due))} ({status_html(task)})"
        for task in tasks
    ]
    return html_text(title) + "\n" + "\n\n".join(task_lines)


def render_task_interval(interval: DateInterval, tasks: list[Task]) -> str:
    return render_tasks(interval.title, select_tasks_for_interval(tasks, interval))


def render_task_list(tasks: list[Task]) -> str:
    if not tasks:
        return "Список задач пуст."
    lines = ["Список задач"]
    task_lines = [
        f"{task.row}. {html_text(task.name)} — {bold(format_date(task.due))} ({status_html(task)})"
        for task in sorted(tasks, key=lambda item: (item.due is None, item.due or date.max, item.row))
    ]
    return "\n".join(lines) + "\n\n" + "\n\n".join(task_lines)


def render_added_tasks(tasks: list[Task]) -> str:
    if not tasks:
        return "Не добавил ни одной задачи."
    title = "Добавил задачу:" if len(tasks) == 1 else "Добавил задачи:"
    task_lines = [
        f"- {html_text(task.name)} — {bold(format_date(task.due))}"
        for task in tasks
    ]
    return title + "\n" + "\n\n".join(task_lines)


def render_delete_confirmation(tasks: list[Task]) -> str:
    lines = ["Вы уверены, что хотите удалить эти задачи?"]
    task_lines = [
        f"- {task.row}. {html_text(task.name)} — {bold(format_date(task.due))} ({status_html(task)})"
        for task in tasks
    ]
    return "\n".join(lines) + "\n\n" + "\n\n".join(task_lines) + "\n\nОтветьте да или нет."


DELETE_CONTEXT_STOP_WORDS = {
    "удали",
    "удалить",
    "удаляй",
    "убери",
    "убрать",
    "закрой",
    "закрыть",
    "заверши",
    "завершить",
    "отметь",
    "отметить",
    "выполнено",
    "выполненной",
    "выполненными",
    "готово",
    "сделано",
    "сделай",
    "перенеси",
    "перенести",
    "перенос",
    "перенесу",
    "сдвинь",
    "сдвинуть",
    "измени",
    "изменить",
    "поменяй",
    "поменять",
    "переименуй",
    "переименовать",
    "назови",
    "назвать",
    "название",
    "наименование",
    "исправь",
    "исправить",
    "запиши",
    "записать",
    "таким",
    "образом",
    "срок",
    "сроки",
    "дату",
    "дата",
    "новую",
    "новая",
    "задача",
    "задачу",
    "задачку",
    "задачи",
    "строка",
    "строку",
    "номер",
    "номера",
    "номером",
    "под",
    "из",
    "списка",
    "про",
    "по",
    "насчет",
    "насчёт",
    "касательно",
    "сейчас",
    "посмотрим",
    "пожалуйста",
    "плиз",
    "мне",
    "надо",
    "нужно",
    "можно",
    "эту",
    "этот",
    "эти",
    "ту",
    "то",
    "с",
    "со",
    "у",
    "о",
    "об",
    "на",
    "в",
    "к",
    "для",
    "и",
    "или",
    "а",
    "же",
    "там",
}


def context_tokens(text: str) -> list[str]:
    cleaned = strip_date_mentions(text).lower().replace("ё", "е")
    words = re.findall(r"[а-яa-z0-9]+", cleaned)
    return [
        word
        for word in words
        if len(word) > 1
        and word not in DELETE_CONTEXT_STOP_WORDS
        and word not in RU_NUMBER_WORDS
    ]


def fuzzy_token_score(query_tokens: list[str], task_tokens: list[str]) -> float:
    if not query_tokens or not task_tokens:
        return 0.0

    best_scores: list[float] = []
    matched = 0
    task_token_set = set(task_tokens)
    exact_matches = len(set(query_tokens) & task_token_set)

    for query_token in query_tokens:
        best = max(SequenceMatcher(None, query_token, task_token).ratio() for task_token in task_tokens)
        best_scores.append(best)
        if best >= 0.74:
            matched += 1

    coverage = matched / len(query_tokens)
    average_similarity = sum(best_scores) / len(best_scores)
    exact_coverage = exact_matches / len(set(query_tokens))
    return coverage * 0.55 + average_similarity * 0.35 + exact_coverage * 0.10


def rank_tasks_by_context(query_tokens: list[str], tasks: list[Task], query_date: date | None = None) -> list[tuple[float, Task]]:
    scored: list[tuple[float, Task]] = []
    for task in tasks:
        task_tokens = context_tokens(task.name)
        score = fuzzy_token_score(query_tokens, task_tokens)
        if query_date is not None:
            score += 0.12 if task.due == query_date else -0.05
        if task.completed:
            score -= 0.15
        if score >= 0.58:
            scored.append((score, task))
    return sorted(scored, key=lambda item: (-item[0], item[1].due or date.max, item[1].row))


def confident_top_match(scored: list[tuple[float, Task]]) -> Task | None:
    if not scored:
        return None

    top_score, top_task = scored[0]
    second_score = scored[1][0] if len(scored) > 1 else 0.0
    if top_score >= 0.72 and top_score - second_score >= 0.10:
        return top_task
    if top_score >= 0.66 and (len(scored) == 1 or top_score - second_score >= 0.08):
        return top_task
    return None


def split_context_fragments(text: str) -> list[list[str]]:
    cleaned = strip_date_mentions(text).lower().replace("ё", "е")
    cleaned = re.sub(r"\b(?:а также|также|плюс|ещ[её])\b", " и ", cleaned)
    cleaned = re.sub(r"[,;/+\n]+", " и ", cleaned)
    fragments: list[list[str]] = []
    for part in re.split(r"\b(?:и|или)\b", cleaned):
        tokens = context_tokens(part)
        if tokens:
            fragments.append(tokens)
    return fragments


def find_multiple_tasks_by_context(text: str, tasks: list[Task], use_date: bool = True) -> tuple[list[Task], list[Task]]:
    candidates = [task for task in tasks if task.name.strip()]
    query_date = extract_query_date(text) if use_date else None
    fragments = split_context_fragments(text)
    if len(fragments) < 2:
        tokens = context_tokens(text)
        if len(tokens) < 2:
            return [], []
        fragments = [[token] for token in tokens if len(token) >= 3]

    matched: list[Task] = []
    suggestions: list[Task] = []
    used_rows: set[int] = set()

    for fragment_tokens in fragments:
        scored = rank_tasks_by_context(fragment_tokens, candidates, query_date)
        top_task = confident_top_match(scored)
        if top_task is None:
            suggestions.extend(task for score, task in scored[:3] if score >= 0.58 and task.row not in used_rows)
            continue
        if top_task.row not in used_rows:
            matched.append(top_task)
            used_rows.add(top_task.row)

    if len(matched) >= 2:
        return matched, []
    if suggestions:
        return [], unique_tasks(suggestions)
    return [], []


def unique_tasks(tasks: list[Task]) -> list[Task]:
    result: list[Task] = []
    seen: set[int] = set()
    for task in tasks:
        if task.row not in seen:
            result.append(task)
            seen.add(task.row)
    return result


def find_tasks_by_delete_context(text: str, tasks: list[Task], use_date: bool = True) -> tuple[list[Task], list[Task]]:
    candidates = [task for task in tasks if task.name.strip()]
    multiple_matches, multiple_suggestions = find_multiple_tasks_by_context(text, candidates, use_date)
    if multiple_matches or multiple_suggestions:
        return multiple_matches, multiple_suggestions

    query_tokens = context_tokens(text)
    query_date = extract_query_date(text) if use_date else None

    if not query_tokens and query_date is not None:
        date_matches = [task for task in candidates if task.due == query_date and not task.completed]
        if len(date_matches) == 1:
            return date_matches, []
        if len(date_matches) > 1:
            return [], date_matches[:5]
        return [], []

    if not query_tokens:
        return [], []

    scored = rank_tasks_by_context(query_tokens, candidates, query_date)
    top_task = confident_top_match(scored)
    if top_task is not None:
        return [top_task], []

    suggestions = [task for score, task in scored[:5] if score >= 0.58]
    return [], suggestions


def render_delete_context_suggestions(tasks: list[Task]) -> str:
    task_lines = [
        f"- {task.row}. {html_text(task.name)} — {bold(format_date(task.due))} ({status_html(task)})"
        for task in tasks
    ]
    return "Нашел несколько похожих задач. Уточните, какую удалить, номером из списка:\n\n" + "\n\n".join(task_lines)


def request_delete_by_text(
    text: str,
    planner_file: Path,
    state: dict | None = None,
    chat_id: int | str | None = None,
    tasks: list[Task] | None = None,
) -> str:
    row_numbers = extract_task_numbers(text)
    if row_numbers:
        return request_delete_confirmation(row_numbers, planner_file, state, chat_id)

    current_tasks = tasks if tasks is not None else get_tasks(planner_file)
    matched_tasks, suggested_tasks = find_tasks_by_delete_context(text, current_tasks)
    if matched_tasks:
        return request_delete_confirmation([task.row for task in matched_tasks], planner_file, state, chat_id)
    if suggested_tasks:
        return request_clarification("delete", render_delete_context_suggestions(suggested_tasks), state, chat_id)
    return request_clarification("delete", "Какие номера задач удалить?", state, chat_id)


def render_move_confirmation(tasks: list[Task], new_due: date) -> str:
    target = "эту задачу" if len(tasks) == 1 else "эти задачи"
    lines = [f"Вы уверены, что хотите перенести {target} на {bold(format_date(new_due))}?"]
    task_lines = [
        f"- {task.row}. {html_text(task.name)} — сейчас {bold(format_date(task.due))} ({status_html(task)})"
        for task in tasks
    ]
    return "\n".join(lines) + "\n\n" + "\n\n".join(task_lines) + "\n\nОтветьте да или нет."


def render_move_context_suggestions(tasks: list[Task], new_due: date) -> str:
    task_lines = [
        f"- {task.row}. {html_text(task.name)} — сейчас {bold(format_date(task.due))} ({status_html(task)})"
        for task in tasks
    ]
    return f"Нашел несколько похожих задач. Уточните номер, какую перенести на {bold(format_date(new_due))}:\n\n" + "\n\n".join(task_lines)


def request_move_confirmation(
    row_numbers: int | list[int],
    new_due: date,
    planner_file: Path,
    state: dict | None = None,
    chat_id: int | str | None = None,
) -> str:
    if isinstance(row_numbers, int):
        rows = [row_numbers]
    else:
        rows = unique_ints([int(row) for row in row_numbers])

    tasks: list[Task] = []
    missing: list[int] = []
    for row_number in rows:
        try:
            tasks.append(get_task_by_row(row_number, planner_file))
        except ValueError:
            missing.append(row_number)

    if not tasks:
        return "Не нашел задачи с номерами: " + ", ".join(map(str, missing))

    if state is not None and chat_id is not None:
        pending = state.setdefault("pending_move", {})
        pending[str(chat_id)] = {
            "rows": [task.row for task in tasks],
            "due_date": new_due.isoformat(),
            "requested_at": datetime.now().isoformat(timespec="seconds"),
        }

    reply = render_move_confirmation(tasks, new_due)
    if missing:
        reply += "\n\nНе нашел задачи с номерами: " + ", ".join(map(str, missing))
    return reply


def request_move_by_text(
    text: str,
    new_due: date,
    planner_file: Path,
    state: dict | None = None,
    chat_id: int | str | None = None,
    tasks: list[Task] | None = None,
) -> str:
    row_numbers = extract_task_numbers(text)
    if row_numbers:
        moved, missing = move_tasks(row_numbers, new_due, planner_file)
        return render_moved_tasks(moved, missing, new_due)

    current_tasks = tasks if tasks is not None else get_tasks(planner_file)
    matched_tasks, suggested_tasks = find_tasks_by_delete_context(text, current_tasks, use_date=False)
    if matched_tasks:
        return request_move_confirmation([task.row for task in matched_tasks], new_due, planner_file, state, chat_id)
    if suggested_tasks:
        return request_clarification(
            "move",
            render_move_context_suggestions(suggested_tasks, new_due),
            state,
            chat_id,
            due_date=new_due.isoformat(),
        )
    return request_clarification(
        "move",
        f"Какие номера задач перенести на {format_date(new_due)}?",
        state,
        chat_id,
        due_date=new_due.isoformat(),
    )


def render_complete_confirmation(tasks: list[Task]) -> str:
    target = "эту задачу" if len(tasks) == 1 else "эти задачи"
    status_word = "выполненной" if len(tasks) == 1 else "выполненными"
    lines = [f"Вы уверены, что хотите отметить {target} {status_word}?"]
    task_lines = [
        f"- {task.row}. {html_text(task.name)} — {bold(format_date(task.due))} ({status_html(task)})"
        for task in tasks
    ]
    return "\n".join(lines) + "\n\n" + "\n\n".join(task_lines) + "\n\nОтветьте да или нет."


def render_complete_context_suggestions(tasks: list[Task]) -> str:
    task_lines = [
        f"- {task.row}. {html_text(task.name)} — {bold(format_date(task.due))} ({status_html(task)})"
        for task in tasks
    ]
    return "Нашел несколько похожих задач. Уточните номер, какие отметить выполненными:\n\n" + "\n\n".join(task_lines)


def request_complete_confirmation(
    row_numbers: int | list[int],
    planner_file: Path,
    state: dict | None = None,
    chat_id: int | str | None = None,
) -> str:
    if isinstance(row_numbers, int):
        rows = [row_numbers]
    else:
        rows = unique_ints([int(row) for row in row_numbers])

    tasks: list[Task] = []
    missing: list[int] = []
    for row_number in rows:
        try:
            tasks.append(get_task_by_row(row_number, planner_file))
        except ValueError:
            missing.append(row_number)

    if not tasks:
        return "Не нашел задачи с номерами: " + ", ".join(map(str, missing))

    if state is not None and chat_id is not None:
        pending = state.setdefault("pending_complete", {})
        pending[str(chat_id)] = {
            "rows": [task.row for task in tasks],
            "requested_at": datetime.now().isoformat(timespec="seconds"),
        }

    reply = render_complete_confirmation(tasks)
    if missing:
        reply += "\n\nНе нашел задачи с номерами: " + ", ".join(map(str, missing))
    return reply


def request_complete_by_text(
    text: str,
    planner_file: Path,
    state: dict | None = None,
    chat_id: int | str | None = None,
    tasks: list[Task] | None = None,
) -> str:
    row_numbers = extract_task_numbers(text)
    if row_numbers:
        completed_tasks, missing = complete_tasks(row_numbers, planner_file)
        return render_completed_tasks(completed_tasks, missing)

    current_tasks = tasks if tasks is not None else get_tasks(planner_file)
    matched_tasks, suggested_tasks = find_tasks_by_delete_context(text, current_tasks, use_date=False)
    if matched_tasks:
        return request_complete_confirmation([task.row for task in matched_tasks], planner_file, state, chat_id)
    if suggested_tasks:
        return request_clarification("complete", render_complete_context_suggestions(suggested_tasks), state, chat_id)
    return request_clarification("complete", "Какие номера задач отметить выполненными?", state, chat_id)


def is_rename_request(text: str) -> bool:
    cleaned = text.lower().replace("ё", "е")
    direct_markers = (
        "переимен",
        "измени название",
        "изменить название",
        "измени наименование",
        "изменить наименование",
        "исправь название",
        "исправить название",
        "исправь наименование",
        "исправить наименование",
        "поменяй название",
        "поменять название",
        "поменяй наименование",
        "поменять наименование",
        "назови задачу",
        "назвать задачу",
    )
    if any(marker in cleaned for marker in direct_markers):
        return True
    if "запиши" in cleaned and extract_task_numbers(cleaned) and (
        any(marker in cleaned for marker in ("таким образом", "следующим образом", "как")) or re.search(r"\bтак\b", cleaned)
    ):
        return True
    if "исправь задачу" in cleaned or "исправить задачу" in cleaned:
        return True
    return False


def clean_new_task_name(value: str) -> str:
    cleaned = value.strip(" -:;,.\"'«»“”")
    cleaned = re.sub(r"^(?:задачу|задача|название|наименование)\b", " ", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned.strip(" -:;,.\"'«»“”")


def extract_new_task_name_from_rename(text: str) -> tuple[str, int] | None:
    cleaned = text.lower().replace("ё", "е")
    patterns = [
        r"\b(?:таким образом|следующим образом|в таком виде)\b\s*[-:—–]?\s*(.+)$",
        r"\b(?:запиши|записать)\b.*?\b(?:как|так)\b\s*[-:—–]?\s*(.+)$",
        r"\b(?:переименуй|переименовать|назови|назвать)\b.*?\b(?:в|на)\b\s*[-:—–]?\s*(.+)$",
        r"\b(?:измени|изменить|исправь|исправить|поменяй|поменять)\b.*?\b(?:на|в)\b\s*[-:—–]?\s*(.+)$",
    ]
    for pattern in patterns:
        match = re.search(pattern, cleaned)
        if not match:
            continue
        start = match.start(1)
        new_name = clean_new_task_name(text[start:])
        if new_name:
            return new_name, start
    return None


def parse_rename_request(text: str) -> tuple[list[int], str] | None:
    if not is_rename_request(text):
        return None
    extracted = extract_new_task_name_from_rename(text)
    if extracted:
        new_name, new_name_start = extracted
        rows = extract_task_numbers(text[:new_name_start]) or extract_task_numbers(text)
        return rows, new_name
    return extract_task_numbers(text), ""


def render_rename_confirmation(tasks: list[Task], new_name: str) -> str:
    target = "этой задачи" if len(tasks) == 1 else "этих задач"
    lines = [f"Вы уверены, что хотите изменить название {target}?"]
    task_lines = [
        "\n".join([
            f"- {task.row}. Было: {html_text(task.name)}",
            f"Будет: {html_text(new_name)}",
            f"Срок: {bold(format_date(task.due))} ({status_html(task)})",
        ])
        for task in tasks
    ]
    return "\n".join(lines) + "\n\n" + "\n\n".join(task_lines) + "\n\nОтветьте да или нет."


def render_renamed_tasks(renamed: list[Task], missing: list[int]) -> str:
    lines: list[str] = []
    if renamed:
        title = "Переименовал задачу:" if len(renamed) == 1 else "Переименовал задачи:"
        lines.append(title)
        lines.append("")
        lines.append("\n\n".join(f"- {task.row}. {html_text(task.name)} — {bold(format_date(task.due))}" for task in renamed))
    if missing:
        if lines:
            lines.append("")
        lines.append("Не нашел задачи с номерами: " + ", ".join(map(str, missing)))
    return "\n".join(lines) if lines else "Не нашел задачи для переименования."


def request_rename_confirmation(
    row_numbers: int | list[int],
    new_name: str,
    planner_file: Path,
    state: dict | None = None,
    chat_id: int | str | None = None,
) -> str:
    if isinstance(row_numbers, int):
        rows = [row_numbers]
    else:
        rows = unique_ints([int(row) for row in row_numbers])

    new_name = clean_new_task_name(new_name)
    if not new_name:
        return request_clarification("rename", "Какое новое название записать для задачи?", state, chat_id, row_numbers=rows)

    tasks: list[Task] = []
    missing: list[int] = []
    for row_number in rows:
        try:
            tasks.append(get_task_by_row(row_number, planner_file))
        except ValueError:
            missing.append(row_number)

    if not tasks:
        return "Не нашел задачи с номерами: " + ", ".join(map(str, missing))

    if state is not None and chat_id is not None:
        pending = state.setdefault("pending_rename", {})
        pending[str(chat_id)] = {
            "rows": [task.row for task in tasks],
            "new_name": new_name,
            "requested_at": datetime.now().isoformat(timespec="seconds"),
        }

    reply = render_rename_confirmation(tasks, new_name)
    if missing:
        reply += "\n\nНе нашел задачи с номерами: " + ", ".join(map(str, missing))
    return reply


def request_rename_by_text(
    text: str,
    new_name: str,
    planner_file: Path,
    state: dict | None = None,
    chat_id: int | str | None = None,
    tasks: list[Task] | None = None,
) -> str:
    row_numbers = extract_task_numbers(text)
    if row_numbers:
        return request_rename_confirmation(row_numbers, new_name, planner_file, state, chat_id)

    current_tasks = tasks if tasks is not None else get_tasks(planner_file)
    matched_tasks, suggested_tasks = find_tasks_by_delete_context(text, current_tasks, use_date=False)
    if matched_tasks:
        return request_rename_confirmation([task.row for task in matched_tasks], new_name, planner_file, state, chat_id)
    if suggested_tasks:
        task_lines = [
            f"- {task.row}. {html_text(task.name)} — {bold(format_date(task.due))} ({status_html(task)})"
            for task in suggested_tasks
        ]
        return request_clarification(
            "rename",
            "Нашел несколько похожих задач. Уточните номер, какую переименовать:\n\n" + "\n\n".join(task_lines),
            state,
            chat_id,
            new_task_name=new_name,
        )
    return request_clarification("rename", "Какой номер задачи переименовать?", state, chat_id, new_task_name=new_name)


def render_moved_tasks(moved: list[Task], missing: list[int], new_due: date) -> str:
    lines: list[str] = []
    if moved:
        lines.append(f"Перенес задачи на {bold(format_date(new_due))}:")
        lines.append("")
        lines.append("\n\n".join(f"- {task.row}. {html_text(task.name)} — {bold(format_date(new_due))}" for task in moved))
    if missing:
        if lines:
            lines.append("")
        lines.append("Не нашел задачи с номерами: " + ", ".join(map(str, missing)))
    return "\n".join(lines) if lines else "Не нашел задачи для переноса."


def render_completed_tasks(completed_tasks: list[Task], missing: list[int]) -> str:
    lines: list[str] = []
    if completed_tasks:
        lines.append("Отметил как выполненные:")
        lines.append("")
        lines.append("\n\n".join(f"- {task.row}. {html_text(task.name)} — {bold(format_date(task.due))}" for task in completed_tasks))
    if missing:
        if lines:
            lines.append("")
        lines.append("Не нашел задачи с номерами: " + ", ".join(map(str, missing)))
    return "\n".join(lines) if lines else "Не нашел задачи для отметки выполнения."


def render_summary(planner_file: Path, soon_days: int) -> str:
    tasks = get_tasks(planner_file)
    today_tasks = select_tasks("today", tasks, soon_days)
    overdue = select_tasks("overdue", tasks, soon_days)
    soon = select_tasks("soon", tasks, soon_days)
    nearest = select_tasks("nearest", tasks, soon_days)

    blocks = [
        render_tasks("Задачи на сегодня", today_tasks),
        render_tasks("Просроченные задачи", overdue),
        render_tasks(f"Скоро, ближайшие {soon_days} дней", soon),
        render_tasks("Ближайшие задачи", nearest),
    ]
    return "\n\n".join(blocks)


def tasks_snapshot_for_ai(tasks: list[Task]) -> str:
    if not tasks:
        return "Таблица задач пуста."

    lines = [
        f"Сегодня: {date.today().isoformat()}",
        "Колонки: номер строки, задача, срок, статус, выполнено, дата выполнения.",
        "",
        "Задачи:",
    ]
    for task in sorted(tasks, key=lambda item: (item.due is None, item.due or date.max, item.row)):
        lines.append(
            " | ".join([
                f"#{task.row}",
                f"задача: {task.name}",
                f"срок: {format_date(task.due)}",
                f"статус: {status_for(task)}",
                f"выполнено: {'да' if task.completed else 'нет'}",
                f"дата выполнения: {format_date(task.completed_on)}",
            ])
        )
    return "\n".join(lines)


def answer_table_question(text: str, planner_file: Path, config: dict) -> str:
    tasks = get_tasks(planner_file)
    snapshot = tasks_snapshot_for_ai(tasks)
    instructions = """
Ты отвечаешь на вопросы пользователя по его Excel-таблице задач.
У тебя есть полный снимок таблицы задач ниже. Используй только эти данные.

Правила:
- Отвечай по-русски, кратко и по делу.
- Можно анализировать, группировать, считать, сравнивать сроки и объяснять приоритеты.
- Нельзя утверждать, что ты изменил таблицу. Изменения делает только бот через отдельные команды.
- Если пользователь просит изменить задачу, скажи, что это можно сделать отдельной командой, например "перенеси задачу 4 на 15 мая" или "отметь задачу 4 выполненной".
- Не используй HTML или Markdown-разметку.
""".strip()
    user_text = f"{snapshot}\n\nВопрос пользователя:\n{text}"
    return html_text(openai_text_response(config, instructions, user_text))


def parse_add_command(text: str) -> tuple[str, date]:
    body = text.strip()
    for prefix in ("/add", "add", "добавить задачу", "добавь задачу", "добавить:", "добавить", "добавь", "+"):
        if body.lower().startswith(prefix):
            body = body[len(prefix):].strip()
            break

    separators = ["|", ";", " до ", " срок ", " к ", " на "]
    for separator in separators:
        if separator in body:
            name, due_text = body.rsplit(separator, 1)
            name = name.strip(" -:;")
            due = parse_date(due_text.strip())
            if not name:
                raise ValueError("Не вижу название задачи.")
            return name, due
    raise ValueError("Формат добавления: /add Название задачи | 25.05.2026")


def is_add_request(text: str) -> bool:
    cleaned = text.lower().replace("ё", "е").strip()
    markers = (
        "/add",
        "add",
        "добав",
        "напом",
        "запиши",
        "создай",
        "поставь",
        "запланируй",
        "будет",
        "запланируй",
        "будет",
    )
    return any(marker in cleaned for marker in markers)


def clean_task_name_from_add_text(text: str) -> str:
    cleaned = strip_date_mentions_preserve_case(text)
    cleaned = re.sub(r"^/add\b", " ", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\b(add|добавь|добавить|добавь задачу|добавить задачу|напомни|запиши|создай|поставь|запланируй|будет)\b", " ", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\b(задачу|задача|срок|дедлайн)\b", " ", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\b(на|в|к|до)\s*(?=$|\bсо\b|\bс\b)", " ", cleaned, flags=re.IGNORECASE)
    cleaned = cleaned.strip()
    cleaned = re.sub(r"^(?:одну|одна|один|1|новую|новая|новый|новое)\b", " ", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"^(?:и|а|также|плюс)\b", " ", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\b(?:и|а|также|плюс)$", " ", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned.strip(" -—–:;,.")


def add_date_mentions(text: str, today: date | None = None) -> list[tuple[int, int, date]]:
    today = today or date.today()
    cleaned = text.lower().replace("ё", "е")
    month_names = "|".join(RU_MONTHS)
    day_words = "|".join(sorted(map(re.escape, RU_DAY_WORDS), key=len, reverse=True))
    weekday_words = "|".join(sorted(map(re.escape, RU_WEEKDAYS), key=len, reverse=True))
    week_context = WEEK_CONTEXT_RE
    patterns = [
        r"\b\d{1,2}[.\/-]\d{1,2}[.\/-]\d{2,4}\b",
        r"\b\d{4}-\d{1,2}-\d{1,2}\b",
        rf"\b\d{{1,2}}\s+(?:{month_names})(?:\s+\d{{4}})?\b",
        rf"\b(?:{day_words})\s+(?:{month_names})(?:\s+\d{{4}})?\b",
        r"\b(?:послезавтра|завтра|сегодня)\b",
        rf"\b(?:на|в|ко?|до)?\s*(?:(?:эт\w*|текущ\w*)\s+)?(?:{NEXT_WEEKDAY_PREFIX_RE}\s+)?(?:{weekday_words})(?:\s+(?:на\s+)?{week_context})?\b",
        rf"\b(?:на\s+)?{week_context}\s+(?:на|в|ко?|до)?\s*(?:{weekday_words})\b",
    ]

    candidates: list[tuple[int, int, date]] = []
    for pattern in patterns:
        for match in re.finditer(pattern, cleaned):
            raw = match.group(0).strip()
            due = extract_query_date(raw, today)
            if due is not None:
                candidates.append((match.start(), match.end(), due))

    result: list[tuple[int, int, date]] = []
    for start, end, due in sorted(candidates, key=lambda item: (item[0], -(item[1] - item[0]))):
        if any(start < accepted_end and end > accepted_start for accepted_start, accepted_end, _ in result):
            continue
        result.append((start, end, due))
    return result


def single_local_date(text: str, today: date | None = None) -> date | None:
    mentions = add_date_mentions(text, today)
    if len(mentions) == 1:
        cleaned = text.lower().replace("ё", "е")
        if re.search(rf"\b{WEEK_CONTEXT_RE}\b", cleaned):
            contextual_due = extract_query_date(text, today)
            if contextual_due is not None:
                return contextual_due
        return mentions[0][2]
    return None


def week_start(day: date) -> date:
    return day - timedelta(days=day.weekday())


def interval_between(start: date, end: date) -> DateInterval:
    if start == end:
        return DateInterval(start, end, f"Задачи на {format_date(start)}")
    return DateInterval(start, end, f"Задачи с {format_date(start)} по {format_date(end)}")


def week_interval(start: date, label: str) -> DateInterval:
    end = start + timedelta(days=6)
    return DateInterval(start, end, f"Задачи на {label} ({format_date(start)} - {format_date(end)})")


def text_week_start(cleaned: str, today: date) -> tuple[date, str] | None:
    current_week = week_start(today)
    if re.search(rf"\b{NEXT_WEEK_CONTEXT_RE}\b", cleaned):
        return current_week + timedelta(days=7), "следующей неделе"
    if re.search(rf"\b{CURRENT_WEEK_CONTEXT_RE}\b", cleaned):
        return current_week, "этой неделе"
    return None


def weekday_pattern() -> str:
    return "|".join(sorted(map(re.escape, RU_WEEKDAYS), key=len, reverse=True))


def date_interval_from_weekday_range(cleaned: str, today: date, week_context: tuple[date, str] | None) -> DateInterval | None:
    match = re.search(rf"\b(?:с|со|от)\s+({weekday_pattern()})\b.*?\b(?:по|до)\s+({weekday_pattern()})\b", cleaned)
    if not match:
        return None
    start_week = week_context[0] if week_context else week_start(today)
    start = start_week + timedelta(days=RU_WEEKDAYS[match.group(1)])
    end = start_week + timedelta(days=RU_WEEKDAYS[match.group(2)])
    if end < start:
        end += timedelta(days=7)
    return interval_between(start, end)


def date_interval_from_date_mentions(text: str, today: date) -> DateInterval | None:
    cleaned = text.lower().replace("ё", "е")
    if not (re.search(r"\b(?:с|со|от)\b", cleaned) and re.search(r"\b(?:по|до)\b", cleaned)):
        return None
    mentions = add_date_mentions(text, today)
    if len(mentions) < 2:
        return None
    start = mentions[0][2]
    end = mentions[1][2]
    if end < start:
        start, end = end, start
    return interval_between(start, end)


def date_interval_after(cleaned: str, text: str, today: date, week_context: tuple[date, str] | None) -> DateInterval | None:
    match = re.search(rf"\bпосле\s+({weekday_pattern()})\b", cleaned)
    if match:
        if week_context:
            anchor = week_context[0] + timedelta(days=RU_WEEKDAYS[match.group(1)])
            end = week_context[0] + timedelta(days=6)
        else:
            anchor = extract_query_date(match.group(1), today)
            end = date.max
        if anchor is None:
            return None
        start = anchor + timedelta(days=1)
        title = f"Задачи после {format_date(anchor)}"
        if end != date.max:
            title += f" по {format_date(end)}"
        return DateInterval(start, end, title)

    if "после" not in cleaned:
        return None
    mentions = add_date_mentions(text, today)
    if len(mentions) != 1:
        return None
    anchor = mentions[0][2]
    start = anchor + timedelta(days=1)
    end = week_context[0] + timedelta(days=6) if week_context else date.max
    title = f"Задачи после {format_date(anchor)}"
    if end != date.max:
        title += f" по {format_date(end)}"
    return DateInterval(start, end, title)


def date_interval_before(cleaned: str, text: str, today: date, week_context: tuple[date, str] | None) -> DateInterval | None:
    match = re.search(rf"\b(?:до|по)\s+({weekday_pattern()})\b", cleaned)
    if match:
        if week_context:
            start = week_context[0]
            anchor = week_context[0] + timedelta(days=RU_WEEKDAYS[match.group(1)])
        else:
            start = today
            anchor = extract_query_date(match.group(1), today)
        if anchor is None:
            return None
        return interval_between(start, anchor)

    if "до" not in cleaned:
        return None
    mentions = add_date_mentions(text, today)
    if len(mentions) != 1:
        return None
    start = week_context[0] if week_context else today
    return interval_between(start, mentions[0][2])


def extract_date_interval(text: str, today: date | None = None) -> DateInterval | None:
    today = today or date.today()
    cleaned = text.lower().replace("ё", "е")
    week_context = text_week_start(cleaned, today)

    interval = date_interval_from_weekday_range(cleaned, today, week_context)
    if interval:
        return interval

    interval = date_interval_from_date_mentions(text, today)
    if interval:
        return interval

    interval = date_interval_after(cleaned, text, today, week_context)
    if interval:
        return interval

    interval = date_interval_before(cleaned, text, today, week_context)
    if interval:
        return interval

    if week_context and "недел" in cleaned:
        return week_interval(week_context[0], week_context[1])

    return None


def parse_multiple_freeform_add(text: str, require_add_request: bool = True) -> list[tuple[str, date]]:
    if require_add_request and not is_add_request(text):
        return []

    mentions = add_date_mentions(text)
    if len(mentions) < 2:
        return []

    parsed: list[tuple[str, date]] = []
    previous_end = 0
    for start, end, due in mentions:
        name = clean_task_name_from_add_text(text[previous_end:start])
        if not name:
            return []
        parsed.append((name, due))
        previous_end = end

    return parsed if len(parsed) >= 2 else []


def parse_freeform_add(text: str) -> tuple[str, date | None] | None:
    if not is_add_request(text):
        return None
    name = clean_task_name_from_add_text(text)
    due = extract_query_date(text)
    if not name:
        return None
    return name, due


def local_adds_from_text(text: str) -> list[tuple[str, date]]:
    parsed_many = parse_multiple_freeform_add(text)
    if parsed_many:
        return parsed_many
    parsed_add = parse_freeform_add(text)
    if parsed_add and parsed_add[1] is not None:
        return [(parsed_add[0], parsed_add[1])]
    return []


def load_config(path: Path = DEFAULT_CONFIG) -> dict:
    # The bot works without an LLM for basic commands.
    # For AI parsing and voice transcription, use any OpenAI-compatible API.
    # Legacy openai_* keys are still accepted for backward compatibility.
    allowed_chat_ids_env = os.getenv("TELEGRAM_ALLOWED_CHAT_IDS", "")
    config = {
        "bot_token": os.getenv("TELEGRAM_BOT_TOKEN", ""),
        "allowed_chat_ids": [item.strip() for item in allowed_chat_ids_env.split(",") if item.strip()],
        "notification_chat_id": os.getenv("TELEGRAM_CHAT_ID", ""),
        "planner_file": str(DEFAULT_PLANNER),
        "soon_days": 7,
        "daily_digest_time": "09:00",
        "llm_api_key": os.getenv("LLM_API_KEY", os.getenv("OPENAI_API_KEY", "")),
        "llm_base_url": os.getenv("LLM_BASE_URL", "https://api.openai.com/v1"),
        "llm_transcription_model": os.getenv("LLM_TRANSCRIPTION_MODEL", "gpt-4o-mini-transcribe"),
        "llm_intent_model": os.getenv("LLM_INTENT_MODEL", "gpt-4o-mini"),
        "openai_api_key": os.getenv("OPENAI_API_KEY", ""),
        "openai_transcription_model": "gpt-4o-mini-transcribe",
        "openai_intent_model": "gpt-4o-mini",
        "voice_language": "ru",
    }
    if path.exists():
        with path.open("r", encoding="utf-8") as file:
            config.update(json.load(file))
    config["planner_file"] = str((BASE_DIR / config["planner_file"]).resolve() if not Path(config["planner_file"]).is_absolute() else Path(config["planner_file"]))
    return config


def api_request(token: str, method: str, payload: dict | None = None) -> dict:
    payload = payload or {}
    data = urllib.parse.urlencode(payload).encode("utf-8")
    url = f"https://api.telegram.org/bot{token}/{method}"
    with urllib.request.urlopen(url, data=data, timeout=70) as response:
        return json.loads(response.read().decode("utf-8"))


def send_message(token: str, chat_id: str | int, text: str) -> None:
    for start in range(0, len(text), 3800):
        api_request(token, "sendMessage", {
            "chat_id": chat_id,
            "text": text[start:start + 3800],
            "parse_mode": "HTML",
            "disable_web_page_preview": "true",
        })


def get_update_message(update: dict) -> tuple[int | None, str, dict]:
    message = update.get("message") or update.get("channel_post") or {}
    chat = message.get("chat") or {}
    return chat.get("id"), (message.get("text") or "").strip(), message


def get_update_text(update: dict) -> tuple[int | None, str]:
    chat_id, text, _ = get_update_message(update)
    return chat_id, text


def download_url(url: str, target: Path) -> None:
    with urllib.request.urlopen(url, timeout=90) as response:
        target.write_bytes(response.read())


def llm_api_key(config: dict) -> str:
    return config.get("llm_api_key") or config.get("openai_api_key", "")


def llm_base_url(config: dict) -> str:
    return (config.get("llm_base_url") or config.get("openai_base_url") or "https://api.openai.com/v1").rstrip("/")


def llm_endpoint(config: dict, path: str) -> str:
    return llm_base_url(config) + path


def llm_intent_model(config: dict) -> str:
    return config.get("llm_intent_model") or config.get("openai_intent_model", "gpt-4o-mini")


def llm_transcription_model(config: dict) -> str:
    return config.get("llm_transcription_model") or config.get("openai_transcription_model", "gpt-4o-mini-transcribe")


def download_telegram_audio(token: str, file_id: str, target_dir: Path, fallback_suffix: str) -> Path:
    file_info = api_request(token, "getFile", {"file_id": file_id})
    file_path = file_info["result"]["file_path"]
    suffix = Path(file_path).suffix or fallback_suffix
    target = target_dir / f"telegram_voice{suffix}"
    download_url(f"https://api.telegram.org/file/bot{token}/{file_path}", target)
    return target


def convert_audio_for_transcription(source: Path) -> Path:
    if source.suffix.lower() in SUPPORTED_TRANSCRIPTION_SUFFIXES:
        return source

    ffmpeg = shutil.which("ffmpeg")
    if not ffmpeg:
        raise RuntimeError(
            "Для голосовых Telegram нужно установить ffmpeg. "
            "Telegram присылает voice в OGG/Opus, а распознавание принимает mp3/m4a/wav/webm."
        )

    target = source.with_suffix(".wav")
    subprocess.run(
        [ffmpeg, "-y", "-i", str(source), "-ar", "16000", "-ac", "1", str(target)],
        check=True,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    return target


def multipart_body(fields: dict[str, str], file_field: str, file_path: Path) -> tuple[bytes, str]:
    boundary = f"----taskbot{uuid4().hex}"
    chunks: list[bytes] = []
    for name, value in fields.items():
        if value:
            chunks.append(f"--{boundary}\r\n".encode("utf-8"))
            chunks.append(f'Content-Disposition: form-data; name="{name}"\r\n\r\n'.encode("utf-8"))
            chunks.append(str(value).encode("utf-8"))
            chunks.append(b"\r\n")

    chunks.append(f"--{boundary}\r\n".encode("utf-8"))
    chunks.append(
        f'Content-Disposition: form-data; name="{file_field}"; filename="{file_path.name}"\r\n'.encode("utf-8")
    )
    chunks.append(b"Content-Type: application/octet-stream\r\n\r\n")
    chunks.append(file_path.read_bytes())
    chunks.append(b"\r\n")
    chunks.append(f"--{boundary}--\r\n".encode("utf-8"))
    return b"".join(chunks), boundary


def transcribe_audio(audio_path: Path, config: dict) -> str:
    api_key = llm_api_key(config)
    if not api_key:
        raise RuntimeError("В telegram_config.json не указан llm_api_key для распознавания голоса.")

    fields = {
        "model": llm_transcription_model(config),
        "language": config.get("voice_language", "ru"),
    }
    body, boundary = multipart_body(fields, "file", audio_path)
    request = urllib.request.Request(
        llm_endpoint(config, "/audio/transcriptions"),
        data=body,
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": f"multipart/form-data; boundary={boundary}",
        },
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=120) as response:
        result = json.loads(response.read().decode("utf-8"))
    return (result.get("text") or "").strip()


def extract_response_text(response: dict) -> str:
    if response.get("output_text"):
        return str(response["output_text"]).strip()
    texts: list[str] = []
    for item in response.get("output", []):
        for content in item.get("content", []):
            if "text" in content:
                texts.append(str(content["text"]))
    return "".join(texts).strip()


def openai_json_response(config: dict, instructions: str, user_text: str, schema: dict) -> dict:
    api_key = llm_api_key(config)
    if not api_key:
        raise RuntimeError("Для понимания свободного текста укажите llm_api_key в telegram_config.json.")

    payload = {
        "model": llm_intent_model(config),
        "instructions": instructions,
        "input": user_text,
        "text": {
            "format": {
                "type": "json_schema",
                "name": "task_intent",
                "strict": True,
                "schema": schema,
            }
        },
    }
    request = urllib.request.Request(
        llm_endpoint(config, "/responses"),
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=90) as response:
        result = json.loads(response.read().decode("utf-8"))
    output_text = extract_response_text(result)
    if not output_text:
        raise RuntimeError("Модель не вернула понятный ответ.")
    return json.loads(output_text)


def openai_text_response(config: dict, instructions: str, user_text: str) -> str:
    api_key = llm_api_key(config)
    if not api_key:
        raise RuntimeError("Для ответов AI по таблице укажите llm_api_key в telegram_config.json.")

    payload = {
        "model": llm_intent_model(config),
        "instructions": instructions,
        "input": user_text,
    }
    request = urllib.request.Request(
        llm_endpoint(config, "/responses"),
        data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=90) as response:
        result = json.loads(response.read().decode("utf-8"))
    output_text = extract_response_text(result)
    if not output_text:
        raise RuntimeError("Модель не вернула понятный ответ.")
    return output_text.strip()


def understand_intent(text: str, config: dict) -> dict:
    schema = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "action": {
                "type": "string",
                "enum": ["add", "delete", "move", "rename", "complete", "list", "today", "date", "range", "overdue", "soon", "nearest", "summary", "help", "unknown"],
            },
            "task_name": {"type": "string"},
            "new_task_name": {"type": "string"},
            "due_date": {"type": "string"},
            "query_date": {"type": "string"},
            "date_from": {"type": "string"},
            "date_to": {"type": "string"},
            "row_number": {"type": "integer"},
            "row_numbers": {
                "type": "array",
                "items": {"type": "integer"},
            },
            "reply_hint": {"type": "string"},
        },
        "required": ["action", "task_name", "new_task_name", "due_date", "query_date", "date_from", "date_to", "row_number", "row_numbers", "reply_hint"],
    }
    instructions = f"""
Ты преобразуешь сообщение пользователя в действие для личного планировщика задач.
Сегодня: {date.today().isoformat()}.
Верни только JSON по схеме.

Действия:
- add: добавить задачу. Заполни task_name и due_date в формате YYYY-MM-DD.
- delete: удалить задачу. Если указан номер, заполни row_number и row_numbers. Если номера нет, но есть описание задачи, заполни task_name.
- move: перенести задачу или задачи на другую дату. Если указаны номера, заполни due_date и row_numbers. Если номеров нет, но есть описание задачи, заполни due_date и task_name.
- rename: изменить наименование существующей задачи. Если указан номер, заполни row_number и row_numbers. Новое название положи в new_task_name.
- complete: отметить задачу или задачи как выполненные. Если указаны номера, заполни row_numbers. Если номеров нет, но есть описание задачи, заполни task_name.
- list: показать все задачи с номерами.
- today: показать задачи на сегодня.
- date: показать задачи на конкретную дату. Заполни query_date в формате YYYY-MM-DD.
- range: показать задачи за период. Заполни date_from и date_to в формате YYYY-MM-DD.
- overdue: показать просроченные задачи.
- soon: показать задачи на ближайшие дни.
- nearest: показать ближайшие будущие задачи.
- summary: показать общую сводку.
- help: пользователь просит помощь.
- unknown: намерение непонятно или не хватает данных.

Правила:
- Если пользователь просит напомнить, записать, поставить, создать или не забыть что-то, это add.
- Если пользователь пишет "запиши задачу на сегодня: ..." или "добавь задачу: ...", это add; вопросительный знак внутри названия задачи не делает запрос question.
- Для add копируй task_name из сообщения с исходными заглавными буквами и аббревиатурами. Не включай в task_name дату и служебные слова: "на завтра", "до этой пятницы", "на эту пятницу", "одну задачу".
- Если срок относительный, пересчитай его относительно сегодняшней даты.
- Если для add нет понятного срока, поставь action unknown и объясни в reply_hint, какой срок нужен.
- Если пользователь спрашивает "какие задачи на 12.05.2026", "что запланировано на 12 мая", "что у меня завтра", "что у меня на среду", "какие задачи на пятницу", это date, а не range и не soon.
- Если пользователь спрашивает "после среды на этой неделе", "на следующей неделе", "с понедельника по среду", это range, а не list.
- Если пользователь просит "переименуй", "исправь название", "измени наименование", "задачу номер 12 запиши таким образом - ...", это rename, а не add и не move.
- Если delete содержит несколько номеров, заполни row_numbers всеми номерами. Если один номер, заполни row_number и row_numbers с одним элементом.
- Если move или complete содержит несколько номеров, заполни row_numbers всеми номерами.
- Если для move нет новой даты, поставь action unknown и попроси уточнить дату.
- Если для complete нет ни номера строки, ни описания задачи, поставь action unknown и попроси уточнить задачу.
- Если для delete нет ни номера строки, ни описания задачи, поставь action unknown и попроси уточнить задачу.
- Не выдумывай срок, если его нет.
""".strip()
    return openai_json_response(config, instructions, text, schema)


def understand_action_plan(text: str, planner_file: Path, config: dict) -> dict:
    tasks = get_tasks(planner_file)
    snapshot = tasks_snapshot_for_ai(tasks)
    schema = {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "actions": {
                "type": "array",
                "items": {
                    "type": "object",
                    "additionalProperties": False,
                    "properties": {
                        "action": {
                            "type": "string",
                            "enum": ["add", "delete", "move", "rename", "complete", "list", "today", "date", "range", "overdue", "soon", "nearest", "summary", "help", "question", "unknown"],
                        },
                        "task_name": {"type": "string"},
                        "new_task_name": {"type": "string"},
                        "due_date": {"type": "string"},
                        "query_date": {"type": "string"},
                        "date_from": {"type": "string"},
                        "date_to": {"type": "string"},
                        "row_numbers": {
                            "type": "array",
                            "items": {"type": "integer"},
                        },
                        "reply_hint": {"type": "string"},
                    },
                    "required": ["action", "task_name", "new_task_name", "due_date", "query_date", "date_from", "date_to", "row_numbers", "reply_hint"],
                },
            },
            "reply_hint": {"type": "string"},
        },
        "required": ["actions", "reply_hint"],
    }
    instructions = f"""
Ты первый слой понимания для личного Telegram-бота задач. Твоя задача — не отвечать пользователю, а перевести его сообщение в строгий JSON-план для скрипта.
Сегодня: {date.today().isoformat()}.
Верни только JSON по схеме.

Доступные действия:
- add: добавить новую задачу. Для каждой новой задачи создай отдельный action add с task_name и due_date YYYY-MM-DD.
- delete: удалить существующую задачу или несколько задач. Используй row_numbers из снимка таблицы.
- move: перенести существующую задачу или несколько задач на новую дату. Используй row_numbers и due_date YYYY-MM-DD.
- rename: изменить наименование существующей задачи, не меняя дату. Используй row_numbers и new_task_name.
- complete: отметить существующую задачу или несколько задач выполненными. Используй row_numbers.
- list, today, overdue, soon, nearest, summary, help: показать соответствующий список или справку.
- date: показать задачи на конкретную дату. Заполни query_date YYYY-MM-DD.
- range: показать задачи за период. Заполни date_from и date_to в формате YYYY-MM-DD.
- question: вопрос по таблице, который не меняет задачи.
- unknown: не хватает данных или намерение непонятно.

Правила:
- Если пользователь просит добавить несколько задач в одном сообщении, верни несколько actions add. Например "добавь плавание на 12 мая и бег на 14 мая" — это два add.
- Если пользователь пишет "запиши задачу на сегодня: ..." или "добавь задачу: ...", это add; вопросительный знак внутри названия задачи не делает запрос question.
- Для add копируй task_name из сообщения с исходными заглавными буквами и аббревиатурами. Не включай в task_name дату и служебные слова: "на завтра", "до этой пятницы", "на эту пятницу", "одну задачу".
- Если пользователь спрашивает "что у меня на среду", "какие задачи на пятницу", "что запланировано на завтра", это date, а не range и не soon.
- Если пользователь просит показать задачи "после среды на этой неделе", "на следующей неделе", "с понедельника по среду", "с 12 мая по 15 мая", это range, а не list.
- Если пользователь просит "переименуй", "исправь название", "измени наименование", "задачу номер 12 запиши таким образом - ...", это rename, а не add и не move.
- Если пользователь просит удалить, перенести или отметить несколько задач, верни один action с row_numbers всех найденных задач.
- Для delete/move/rename/complete сопоставляй описания пользователя со снимком таблицы. Например "бег и плавание" должны стать номерами соответствующих строк, если такие задачи есть.
- Не выдумывай номера строк: используй только строки из снимка таблицы.
- Если совпадение неоднозначно или не найдено, верни unknown и в reply_hint задай короткий уточняющий вопрос.
- Если не хватает даты для add или move, верни unknown и попроси дату.
- Если для rename нет новой формулировки задачи или непонятно, какую строку менять, верни unknown и попроси уточнить.
- Если пользователь задает аналитический вопрос по таблице, верни question.
- Не меняй таблицу сам. Скрипт сам проверит и выполнит действия.
""".strip()
    user_text = f"{snapshot}\n\nСообщение пользователя:\n{text}"
    return openai_json_response(config, instructions, user_text, schema)


def transcribe_telegram_message(token: str, message: dict, config: dict) -> str:
    audio = message.get("voice") or message.get("audio")
    if not audio:
        raise RuntimeError("В сообщении нет голосового файла.")

    file_id = audio["file_id"]
    fallback_suffix = Path(audio.get("file_name", "")).suffix or ".ogg"
    with tempfile.TemporaryDirectory() as tmp:
        source = download_telegram_audio(token, file_id, Path(tmp), fallback_suffix)
        prepared = convert_audio_for_transcription(source)
        return transcribe_audio(prepared, config)


def normalize_voice_command(text: str) -> str:
    original = text.strip()
    cleaned = original.lower()
    cleaned = cleaned.replace("ё", "е").strip(" .,!?:;")

    if cleaned.startswith(("добавь", "добавить")):
        return original
    if is_rename_request(cleaned):
        return original
    if is_move_request(cleaned) or is_complete_request(cleaned):
        return original
    if "удали" in cleaned or "удалить" in cleaned:
        row_numbers = extract_task_numbers(cleaned)
        return "/delete " + " ".join(map(str, row_numbers)) if row_numbers else original
    if extract_date_interval(cleaned) is not None:
        return original
    if "список" in cleaned and extract_query_date(cleaned) is None:
        return "/list"
    if "свод" in cleaned or "статус" in cleaned:
        return "/summary"
    if "просроч" in cleaned:
        return "/overdue"
    if "сегодня" in cleaned:
        return "/today"
    if "ближай" in cleaned:
        return "/nearest"
    if "скоро" in cleaned:
        return "/soon"
    return original


def is_move_request(text: str) -> bool:
    cleaned = text.lower().replace("ё", "е")
    return any(word in cleaned for word in ("перенеси", "перенести", "перенос", "сдвинь", "сдвинуть", "измени срок", "поменяй срок"))


def is_complete_request(text: str) -> bool:
    cleaned = text.lower().replace("ё", "е")
    return any(word in cleaned for word in ("выполн", "готово", "закрыл", "закрыть", "отметь", "сделано", "заверш"))


def _strip_date_mentions(cleaned: str, flags: int = 0) -> str:
    cleaned = re.sub(r"\b\d{1,2}[.\/-]\d{1,2}[.\/-]\d{2,4}\b", " ", cleaned)
    cleaned = re.sub(r"\b\d{4}-\d{1,2}-\d{1,2}\b", " ", cleaned)
    cleaned = re.sub(r"\b(?:на|в|ко?|до)?\s*(?:послезавтра|завтра|сегодня)\b", " ", cleaned, flags=flags)
    month_names = "|".join(RU_MONTHS)
    cleaned = re.sub(rf"\b\d{{1,2}}\s+(?:{month_names})(?:\s+\d{{4}})?\b", " ", cleaned, flags=flags)
    day_words = "|".join(sorted(map(re.escape, RU_DAY_WORDS), key=len, reverse=True))
    cleaned = re.sub(rf"\b(?:{day_words})\s+(?:{month_names})(?:\s+\d{{4}})?\b", " ", cleaned, flags=flags)
    weekday_words = "|".join(sorted(map(re.escape, RU_WEEKDAYS), key=len, reverse=True))
    week_context = WEEK_CONTEXT_RE
    cleaned = re.sub(rf"\b(?:на|в|ко?|до)?\s*(?:(?:эт\w*|текущ\w*)\s+)?(?:{NEXT_WEEKDAY_PREFIX_RE}\s+)?(?:{weekday_words})(?:\s+(?:на\s+)?{week_context})?\b", " ", cleaned, flags=flags)
    cleaned = re.sub(rf"\b(?:на\s+)?{week_context}\s+(?:на|в|ко?|до)?\s*(?:{weekday_words})\b", " ", cleaned, flags=flags)
    cleaned = re.sub(rf"\b(?:{weekday_words})\b", " ", cleaned, flags=flags)
    return cleaned


def strip_date_mentions(text: str) -> str:
    return _strip_date_mentions(text.lower().replace("ё", "е"))


def strip_date_mentions_preserve_case(text: str) -> str:
    return _strip_date_mentions(text, re.IGNORECASE)


def extract_task_number(text: str) -> int | None:
    numbers = extract_task_numbers(text)
    return numbers[-1] if numbers else None


def extract_task_numbers(text: str) -> list[int]:
    text = strip_date_mentions(text)
    digits = re.findall(r"\d+", text)
    if digits:
        return unique_ints([int(item) for item in digits])

    words = re.findall(r"[а-яa-z]+", text.lower().replace("ё", "е"))
    skip_words = {
        "номер",
        "номера",
        "номером",
        "задача",
        "задачу",
        "задачи",
        "строка",
        "строку",
        "под",
        "с",
        "и",
        "пожалуйста",
        "удали",
        "удалить",
    }
    numbers: list[int] = []
    current = 0
    current_has_compound_part = False
    for word in words + [""]:
        if word in RU_NUMBER_WORDS:
            value = RU_NUMBER_WORDS[word]
            if current and not current_has_compound_part and 0 < current < 10 and 0 < value < 10:
                numbers.append(current)
                current = value
            else:
                current += value
            if value >= 10:
                current_has_compound_part = True
            continue
        if current:
            numbers.append(current)
            current = 0
            current_has_compound_part = False
        if word in skip_words:
            continue
    return unique_ints(numbers)


def unique_ints(values: list[int]) -> list[int]:
    result: list[int] = []
    seen: set[int] = set()
    for value in values:
        if value > 0 and value not in seen:
            result.append(value)
            seen.add(value)
    return result


def unique_texts(values) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        cleaned = str(value or "").strip()
        key = cleaned.lower()
        if cleaned and key not in seen:
            result.append(cleaned)
            seen.add(key)
    return result


def parse_yes_no(text: str, config: dict | None = None) -> bool | None:
    cleaned = text.strip().lower().replace("ё", "е")
    cleaned = cleaned.strip(" .,!?:;")
    yes_words = {
        "да",
        "ага",
        "угу",
        "верно",
        "подтверждаю",
        "удаляй",
        "удалить",
        "можно",
        "конечно",
        "yes",
        "y",
        "ok",
        "okay",
    }
    no_words = {
        "нет",
        "не",
        "отмена",
        "отмени",
        "не надо",
        "не удаляй",
        "оставь",
        "no",
        "n",
        "cancel",
    }
    if cleaned in yes_words:
        return True
    if cleaned in no_words:
        return False
    if any(phrase in cleaned for phrase in ("не удаляй", "не надо", "отмени", "оставь")):
        return False
    if any(word in cleaned.split() for word in yes_words):
        return True
    if any(word in cleaned.split() for word in no_words):
        return False

    if config and llm_api_key(config):
        schema = {
            "type": "object",
            "additionalProperties": False,
            "properties": {
                "answer": {"type": "string", "enum": ["yes", "no", "unknown"]},
            },
            "required": ["answer"],
        }
        instructions = "Определи, подтверждает ли пользователь удаление задач. Верни yes, no или unknown."
        try:
            result = openai_json_response(config, instructions, text, schema)
            if result.get("answer") == "yes":
                return True
            if result.get("answer") == "no":
                return False
        except Exception:
            return None
    return None


def allowed(chat_id: int | None, allowed_chat_ids: list) -> bool:
    if chat_id is None:
        return False
    if not allowed_chat_ids:
        return False
    return str(chat_id) in {str(item) for item in allowed_chat_ids}


def help_text(chat_id: int | None = None) -> str:
    intro = "Я умею добавлять задачи в Excel и показывать сводку."
    if chat_id is not None:
        intro += f"\nChat ID: {chat_id}"
    return "\n".join([
        intro,
        "",
        "Команды:",
        "/add Название задачи | 25.05.2026",
        "/add Оплатить счет | завтра",
        "/list — список задач с номерами",
        "/delete 4 — удалить задачу под номером 4",
        "Перенести: перенеси задачу 4 на 15 мая",
        "Переименовать: задачу 4 запиши так - новое название",
        "Выполнить: отметь задачи 2 и 3 выполненными",
        "Голосом: скажите те же команды обычными словами",
        "/today — задачи на сегодня",
        "/overdue — просроченные",
        "/soon — ближайшие дни",
        "/nearest — ближайшие задачи",
        "/summary — полная сводка",
        "/help — помощь",
    ])


def setup_lock_text(chat_id: int | None) -> str:
    return "\n".join([
        "Бот еще не привязан к владельцу.",
        f"Ваш Chat ID: {chat_id}",
        "",
        "Чтобы включить личный доступ, добавьте этот ID в telegram_config.json:",
        f'"allowed_chat_ids": [{chat_id}],',
        f'"notification_chat_id": "{chat_id}"',
        "",
        "После сохранения файла перезапустите бота.",
    ])


def request_delete_confirmation(row_numbers: int | list[int], planner_file: Path, state: dict | None = None, chat_id: int | str | None = None) -> str:
    if isinstance(row_numbers, int):
        rows = [row_numbers]
    else:
        rows = unique_ints([int(row) for row in row_numbers])

    tasks: list[Task] = []
    missing: list[int] = []
    for row_number in rows:
        try:
            tasks.append(get_task_by_row(row_number, planner_file))
        except ValueError:
            missing.append(row_number)

    if not tasks:
        return "Не нашел задачи с номерами: " + ", ".join(map(str, missing))

    if state is not None and chat_id is not None:
        pending = state.setdefault("pending_delete", {})
        pending[str(chat_id)] = {
            "rows": [task.row for task in tasks],
            "requested_at": datetime.now().isoformat(timespec="seconds"),
        }

    reply = render_delete_confirmation(tasks)
    if missing:
        reply += "\n\nНе нашел задачи с номерами: " + ", ".join(map(str, missing))
    return reply


def execute_pending_delete(state: dict, chat_id: int | str, planner_file: Path) -> str:
    pending = state.get("pending_delete", {}).pop(str(chat_id), None)
    if not pending:
        return "Нет ожидающего удаления."

    deleted: list[Task] = []
    missing: list[int] = []
    for row_number in pending.get("rows", []):
        try:
            deleted.append(delete_task(int(row_number), planner_file))
        except ValueError:
            missing.append(int(row_number))

    lines = []
    if deleted:
        lines.append("Удалил задачи:")
        lines.append("")
        lines.append("\n\n".join(f"- {task.row}. {html_text(task.name)} — {bold(format_date(task.due))}" for task in deleted))
    if missing:
        if lines:
            lines.append("")
        lines.append("Не нашел задачи с номерами: " + ", ".join(map(str, missing)))
    return "\n".join(lines) if lines else "Удалять уже нечего."


def cancel_pending_delete(state: dict, chat_id: int | str) -> str:
    state.get("pending_delete", {}).pop(str(chat_id), None)
    return "Ок, удаление отменено."


def execute_pending_move(state: dict, chat_id: int | str, planner_file: Path) -> str:
    pending = state.get("pending_move", {}).pop(str(chat_id), None)
    if not pending:
        return "Нет ожидающего переноса."

    due = pending_due_to_date(pending.get("due_date"))
    if due is None:
        return "Не понял новую дату для переноса."

    row_numbers = unique_ints([int(row) for row in pending.get("rows", []) if int(row) >= 2])
    moved, missing = move_tasks(row_numbers, due, planner_file)
    return render_moved_tasks(moved, missing, due)


def cancel_pending_move(state: dict, chat_id: int | str) -> str:
    state.get("pending_move", {}).pop(str(chat_id), None)
    return "Ок, перенос отменен."


def execute_pending_rename(state: dict, chat_id: int | str, planner_file: Path) -> str:
    pending = state.get("pending_rename", {}).pop(str(chat_id), None)
    if not pending:
        return "Нет ожидающего переименования."

    row_numbers = unique_ints([int(row) for row in pending.get("rows", []) if int(row) >= 2])
    new_name = clean_new_task_name(str(pending.get("new_name") or ""))
    if not new_name:
        return "Не понял новое название задачи."
    renamed, missing = rename_tasks([(row_number, new_name) for row_number in row_numbers], planner_file)
    return render_renamed_tasks(renamed, missing)


def cancel_pending_rename(state: dict, chat_id: int | str) -> str:
    state.get("pending_rename", {}).pop(str(chat_id), None)
    return "Ок, переименование отменено."


def execute_pending_complete(state: dict, chat_id: int | str, planner_file: Path) -> str:
    pending = state.get("pending_complete", {}).pop(str(chat_id), None)
    if not pending:
        return "Нет ожидающей отметки выполнения."

    row_numbers = unique_ints([int(row) for row in pending.get("rows", []) if int(row) >= 2])
    completed_tasks, missing = complete_tasks(row_numbers, planner_file)
    return render_completed_tasks(completed_tasks, missing)


def cancel_pending_complete(state: dict, chat_id: int | str) -> str:
    state.get("pending_complete", {}).pop(str(chat_id), None)
    return "Ок, отметку выполнения отменил."


def is_stop_command(text: str) -> bool:
    cleaned = text.strip().lower().replace("ё", "е")
    cleaned = cleaned.strip(" .,!?:;")
    return cleaned in {
        "стоп",
        "stop",
        "отмена",
        "отмени",
        "сброс",
        "сбрось",
        "не надо",
        "cancel",
    }


def is_cancel_reply(text: str) -> bool:
    cleaned = text.strip().lower().replace("ё", "е")
    cleaned = cleaned.strip(" .,!?:;")
    return is_stop_command(text) or cleaned in {"нет", "не", "no", "nope"}


def is_new_request(text: str) -> bool:
    cleaned = text.strip().lower().replace("ё", "е")
    if not cleaned:
        return False
    if cleaned.startswith(("/", "+")):
        return True
    if is_add_request(cleaned) or is_change_request(cleaned):
        return True
    read_markers = (
        "покажи",
        "показать",
        "какие",
        "какая",
        "какой",
        "что у меня",
        "что на",
        "список",
        "сводка",
        "статус",
        "просроч",
        "ближай",
        "скоро",
    )
    return any(marker in cleaned for marker in read_markers)


def has_pending_state(state: dict, chat_id: int | str) -> bool:
    chat_key = str(chat_id)
    return any(state.get(key, {}).get(chat_key) for key in ("pending_delete", "pending_move", "pending_rename", "pending_complete", "pending_add", "pending_clarification"))


def clear_pending_state(state: dict, chat_id: int | str) -> None:
    chat_key = str(chat_id)
    for key in ("pending_delete", "pending_move", "pending_rename", "pending_complete", "pending_add", "pending_clarification"):
        state.get(key, {}).pop(chat_key, None)


def request_clarification(
    action: str,
    question: str,
    state: dict | None = None,
    chat_id: int | str | None = None,
    **data,
) -> str:
    if state is not None and chat_id is not None:
        pending = state.setdefault("pending_clarification", {})
        pending[str(chat_id)] = {
            "action": action,
            "data": data,
            "requested_at": datetime.now().isoformat(timespec="seconds"),
        }
    return question


def pending_due_to_date(value: str | None) -> date | None:
    if not value:
        return None
    due = extract_query_date(value)
    if due is not None:
        return due
    try:
        return parse_date(value)
    except ValueError:
        return None


def clean_task_name_from_reply(text: str) -> str:
    parsed_add = parse_freeform_add(text)
    if parsed_add:
        return parsed_add[0]
    name = strip_date_mentions(text)
    name = re.sub(r"\b(задача|задачу|название|назови|это|будет)\b", " ", name)
    name = re.sub(r"\s+", " ", name)
    return name.strip(" -:;,.")


def request_add_date(task_name: str, state: dict | None = None, chat_id: int | str | None = None) -> str:
    if state is not None and chat_id is not None:
        pending = state.setdefault("pending_add", {})
        pending[str(chat_id)] = {
            "task_name": task_name,
            "requested_at": datetime.now().isoformat(timespec="seconds"),
        }
    return f"На какую дату добавить задачу: {html_text(task_name)}?"


def handle_pending_add(text: str, state: dict, chat_id: int | str, planner_file: Path) -> str | None:
    pending = state.get("pending_add", {}).get(str(chat_id))
    if not pending:
        return None

    if is_stop_command(text) or text.strip().lower().replace("ё", "е") == "нет":
        state.get("pending_add", {}).pop(str(chat_id), None)
        return "Ок, добавление отменено."

    parsed_many = parse_multiple_freeform_add(text, require_add_request=False)
    if parsed_many:
        state.get("pending_add", {}).pop(str(chat_id), None)
        added = [add_task(name, due, planner_file) for name, due in parsed_many]
        return render_added_tasks(added)

    due = extract_query_date(text)
    if due is None:
        return "Я жду дату для новой задачи. Например: 13 мая, завтра или 13.05.2026."

    state.get("pending_add", {}).pop(str(chat_id), None)
    task = add_task(pending["task_name"], due, planner_file)
    return render_added_tasks([task])


def handle_pending_clarification(
    text: str,
    state: dict,
    chat_id: int | str,
    planner_file: Path,
    soon_days: int,
    config: dict | None = None,
) -> str | None:
    pending = state.get("pending_clarification", {}).get(str(chat_id))
    if not pending:
        return None

    if is_cancel_reply(text):
        state.get("pending_clarification", {}).pop(str(chat_id), None)
        return "Ок, ожидание сброшено."

    action = pending.get("action", "")
    data = pending.get("data", {})

    if action == "add":
        parsed_many = parse_multiple_freeform_add(text, require_add_request=False)
        if parsed_many:
            state.get("pending_clarification", {}).pop(str(chat_id), None)
            added = [add_task(name, due, planner_file) for name, due in parsed_many]
            return render_added_tasks(added)

        task_name = (data.get("task_name") or "").strip()
        due = pending_due_to_date(data.get("due_date"))
        if not task_name:
            task_name = clean_task_name_from_reply(text)
        if due is None:
            due = extract_query_date(text)

        if task_name and due is not None:
            state.get("pending_clarification", {}).pop(str(chat_id), None)
            task = add_task(task_name, due, planner_file)
            return render_added_tasks([task])

        data.update({"task_name": task_name, "due_date": due.isoformat() if due else ""})
        return request_clarification(
            "add",
            "Уточните, пожалуйста, название задачи и срок. Например: созвон со Свеем на 13 мая.",
            state,
            chat_id,
            **data,
        )

    if action == "delete":
        state.get("pending_clarification", {}).pop(str(chat_id), None)
        return request_delete_by_text(text, planner_file, state, chat_id)

    if action == "move":
        row_numbers = unique_ints([int(row) for row in data.get("row_numbers", []) if int(row) >= 2])
        due = pending_due_to_date(data.get("due_date"))
        if not row_numbers:
            row_numbers = extract_task_numbers(text)
        if due is None:
            due = extract_query_date(text)

        if row_numbers and due is not None:
            state.get("pending_clarification", {}).pop(str(chat_id), None)
            moved, missing = move_tasks(row_numbers, due, planner_file)
            return render_moved_tasks(moved, missing, due)
        if not row_numbers and due is not None:
            state.get("pending_clarification", {}).pop(str(chat_id), None)
            return request_move_by_text(text, due, planner_file, state, chat_id)

        data.update({"row_numbers": row_numbers, "due_date": due.isoformat() if due else ""})
        if not row_numbers and due is None:
            question = "Какие номера задач перенести и на какую дату?"
        elif not row_numbers:
            question = f"Какие номера задач перенести на {format_date(due)}?"
        else:
            question = "На какую дату перенести задачи: " + ", ".join(map(str, row_numbers)) + "?"
        return request_clarification("move", question, state, chat_id, **data)

    if action == "rename":
        row_numbers = unique_ints([int(row) for row in data.get("row_numbers", []) if int(row) >= 2])
        new_name = clean_new_task_name(str(data.get("new_task_name") or ""))
        parsed_rename = parse_rename_request(text)
        if parsed_rename:
            parsed_rows, parsed_name = parsed_rename
            if not row_numbers:
                row_numbers = parsed_rows
            if not new_name:
                new_name = parsed_name
        if not row_numbers:
            row_numbers = extract_task_numbers(text)
        if not new_name and row_numbers:
            new_name = clean_new_task_name(text)

        if row_numbers and new_name:
            state.get("pending_clarification", {}).pop(str(chat_id), None)
            return request_rename_confirmation(row_numbers, new_name, planner_file, state, chat_id)

        data.update({"row_numbers": row_numbers, "new_task_name": new_name})
        if not row_numbers and not new_name:
            question = "Какую задачу переименовать и какое новое название записать?"
        elif not row_numbers:
            question = "Какой номер задачи переименовать?"
        else:
            question = "Какое новое название записать для задачи: " + ", ".join(map(str, row_numbers)) + "?"
        return request_clarification("rename", question, state, chat_id, **data)

    if action == "complete":
        row_numbers = unique_ints([int(row) for row in data.get("row_numbers", []) if int(row) >= 2])
        if not row_numbers:
            row_numbers = extract_task_numbers(text)
        if row_numbers:
            state.get("pending_clarification", {}).pop(str(chat_id), None)
            completed_tasks, missing = complete_tasks(row_numbers, planner_file)
            return render_completed_tasks(completed_tasks, missing)
        state.get("pending_clarification", {}).pop(str(chat_id), None)
        return request_complete_by_text(text, planner_file, state, chat_id)

    if action == "ai_unknown":
        original_text = (data.get("original_text") or "").strip()
        state.get("pending_clarification", {}).pop(str(chat_id), None)
        combined_text = f"{original_text}\nУточнение пользователя: {text}" if original_text else text
        if config and llm_api_key(config):
            return handle_ai_intent_or_question(combined_text, planner_file, soon_days, config, state, chat_id)
        return handle_text(combined_text, planner_file, soon_days, config, state, chat_id)

    state.get("pending_clarification", {}).pop(str(chat_id), None)
    return handle_text(text, planner_file, soon_days, config, state, chat_id)


def is_change_request(text: str) -> bool:
    cleaned = text.lower().replace("ё", "е")
    markers = (
        "добав",
        "напом",
        "запиши",
        "создай",
        "поставь",
        "удали",
        "удалить",
        "перенеси",
        "перенести",
        "сдвинь",
        "измени срок",
        "поменяй срок",
        "переимен",
        "название",
        "наименование",
        "назови",
        "исправь",
        "исправить",
        "таким образом",
        "выполн",
        "закрой",
        "закрыть",
        "отметь",
        "сделано",
        "заверш",
    )
    return any(marker in cleaned for marker in markers)


def handle_ai_intent_or_question(
    text: str,
    planner_file: Path,
    soon_days: int,
    config: dict,
    state: dict | None = None,
    chat_id: int | str | None = None,
) -> str:
    intent = understand_intent(text, config)
    if intent.get("action") == "unknown":
        task_name = (intent.get("task_name") or "").strip()
        if task_name and is_add_request(text):
            return request_add_date(task_name, state, chat_id)
        if is_change_request(text):
            return request_clarification(
                "ai_unknown",
                intent.get("reply_hint") or "Не понял, какое изменение нужно сделать. Уточните задачу, номер и дату, если нужна дата.",
                state,
                chat_id,
                original_text=text,
            )
        return answer_table_question(text, planner_file, config)
    return execute_intent(intent, planner_file, soon_days, state, chat_id, original_text=text)


def ai_row_numbers(action: dict) -> list[int]:
    rows: list[int] = []
    for row in action.get("row_numbers", []):
        try:
            number = int(row)
        except (TypeError, ValueError):
            continue
        if number >= 2:
            rows.append(number)
    return unique_ints(rows)


def action_type(action: dict) -> str:
    return str(action.get("action") or "unknown").strip().lower()


def execute_action_plan(
    plan: dict,
    original_text: str,
    planner_file: Path,
    soon_days: int,
    config: dict,
    state: dict | None = None,
    chat_id: int | str | None = None,
) -> str | None:
    actions = [action for action in plan.get("actions", []) if action_type(action) != "unknown"]
    types = {action_type(action) for action in actions}
    tasks = get_tasks(planner_file)
    read_only_types = {"list", "today", "date", "range", "overdue", "soon", "nearest", "summary", "help", "question"}
    local_interval = extract_date_interval(original_text)
    if local_interval is not None and (not actions or types <= read_only_types) and not is_change_request(original_text):
        return render_task_interval(local_interval, tasks)

    local_query_date = single_local_date(original_text)
    date_list_types = {"list", "today", "date", "range", "soon", "nearest"}
    if local_query_date is not None and (not actions or types <= date_list_types) and not is_change_request(original_text):
        return render_tasks(f"Задачи на {format_date(local_query_date)}", select_tasks_for_date(tasks, local_query_date))

    if not actions:
        hint = (plan.get("reply_hint") or "").strip()
        if hint:
            return request_clarification("ai_unknown", hint, state, chat_id, original_text=original_text)
        return None

    if types == {"add"}:
        missing = [action for action in actions if not (action.get("task_name") or "").strip() or not (action.get("due_date") or "").strip()]
        if missing:
            return request_clarification(
                "add",
                "Уточните, пожалуйста, название и срок для каждой новой задачи.",
                state,
                chat_id,
                original_text=original_text,
            )
        added: list[Task] = []
        local_adds = local_adds_from_text(original_text)
        if len(local_adds) == len(actions):
            for task_name, due in local_adds:
                added.append(add_task(task_name, due, planner_file))
            return render_added_tasks(added)

        local_due = single_local_date(original_text) if len(actions) == 1 else None
        for action in actions:
            try:
                due = local_due or parse_date(str(action["due_date"]))
                added.append(add_task(str(action["task_name"]).strip(), due, planner_file))
            except ValueError:
                return request_clarification(
                    "add",
                    "Не понял одну из дат. Напишите, пожалуйста, задачи и сроки ещё раз.",
                    state,
                    chat_id,
                    original_text=original_text,
                )
        return render_added_tasks(added)

    if types == {"delete"}:
        rows = unique_ints([row for action in actions for row in ai_row_numbers(action)])
        if rows:
            return request_delete_confirmation(rows, planner_file, state, chat_id)
        names = " и ".join((action.get("task_name") or "").strip() for action in actions if (action.get("task_name") or "").strip())
        if names:
            return request_delete_by_text(names, planner_file, state, chat_id, tasks)
        return request_clarification("delete", "Какие задачи удалить?", state, chat_id)

    if types == {"move"}:
        due_values = [str(action.get("due_date") or "").strip() for action in actions if str(action.get("due_date") or "").strip()]
        parsed_due: list[date] = []
        local_due = single_local_date(original_text)
        if local_due is not None:
            parsed_due.append(local_due)
        else:
            for due_text in due_values:
                try:
                    due = parse_date(due_text)
                except ValueError:
                    return request_clarification("move", "На какую дату перенести задачи?", state, chat_id, original_text=original_text)
                if due not in parsed_due:
                    parsed_due.append(due)
        if len(parsed_due) != 1:
            return request_clarification("move", "Уточните одну дату для переноса задач.", state, chat_id, original_text=original_text)
        rows = unique_ints([row for action in actions for row in ai_row_numbers(action)])
        new_due = parsed_due[0]
        if rows:
            explicit_rows = extract_task_numbers(original_text)
            if len(rows) > 1 or not explicit_rows:
                return request_move_confirmation(rows, new_due, planner_file, state, chat_id)
            moved, missing = move_tasks(rows, new_due, planner_file)
            return render_moved_tasks(moved, missing, new_due)
        names = " и ".join((action.get("task_name") or "").strip() for action in actions if (action.get("task_name") or "").strip())
        if names:
            return request_move_by_text(names, new_due, planner_file, state, chat_id, tasks)
        return request_clarification("move", f"Какие задачи перенести на {format_date(new_due)}?", state, chat_id, due_date=new_due.isoformat())

    if types == {"rename"}:
        new_names = unique_texts(str(action.get("new_task_name") or "").strip() for action in actions)
        if len(new_names) != 1:
            return request_clarification("rename", "Какое новое название записать для задачи?", state, chat_id, original_text=original_text)
        new_name = new_names[0]
        rows = unique_ints([row for action in actions for row in ai_row_numbers(action)])
        if rows:
            return request_rename_confirmation(rows, new_name, planner_file, state, chat_id)
        names = " и ".join((action.get("task_name") or "").strip() for action in actions if (action.get("task_name") or "").strip())
        if names:
            return request_rename_by_text(names, new_name, planner_file, state, chat_id, tasks)
        return request_clarification("rename", "Какую задачу переименовать?", state, chat_id, new_task_name=new_name)

    if types == {"complete"}:
        rows = unique_ints([row for action in actions for row in ai_row_numbers(action)])
        if rows:
            explicit_rows = extract_task_numbers(original_text)
            if len(rows) > 1 or not explicit_rows:
                return request_complete_confirmation(rows, planner_file, state, chat_id)
            completed_tasks, missing = complete_tasks(rows, planner_file)
            return render_completed_tasks(completed_tasks, missing)
        names = " и ".join((action.get("task_name") or "").strip() for action in actions if (action.get("task_name") or "").strip())
        if names:
            return request_complete_by_text(names, planner_file, state, chat_id, tasks)
        return request_clarification("complete", "Какие задачи отметить выполненными?", state, chat_id)

    if len(actions) == 1:
        action = actions[0]
        kind = action_type(action)
        if kind == "list":
            return render_task_list(tasks)
        if kind == "today":
            return render_tasks("Задачи на сегодня", select_tasks("today", tasks, soon_days))
        if kind == "date":
            query_date = (action.get("query_date") or "").strip()
            local_query_date = single_local_date(original_text)
            if not query_date and local_query_date is None:
                return request_clarification("date", "На какую дату показать задачи?", state, chat_id)
            try:
                target_date = local_query_date or parse_date(query_date)
            except ValueError:
                return request_clarification("date", "Не понял дату. Например: 12.05.2026.", state, chat_id)
            return render_tasks(f"Задачи на {format_date(target_date)}", select_tasks_for_date(tasks, target_date))
        if kind == "range":
            date_from = (action.get("date_from") or "").strip()
            date_to = (action.get("date_to") or "").strip()
            if not date_from or not date_to:
                return request_clarification("range", "За какой период показать задачи?", state, chat_id)
            try:
                interval = interval_between(parse_date(date_from), parse_date(date_to))
            except ValueError:
                return request_clarification("range", "Не понял период. Например: с 12.05.2026 по 15.05.2026.", state, chat_id)
            return render_task_interval(interval, tasks)
        if kind == "overdue":
            return render_tasks("Просроченные задачи", select_tasks("overdue", tasks, soon_days))
        if kind == "soon":
            return render_tasks(f"Скоро, ближайшие {soon_days} дней", select_tasks("soon", tasks, soon_days))
        if kind == "nearest":
            return render_tasks("Ближайшие задачи", select_tasks("nearest", tasks, soon_days))
        if kind == "summary":
            return render_summary(planner_file, soon_days)
        if kind == "help":
            return help_text()
        if kind == "question":
            return answer_table_question(original_text, planner_file, config)

    return request_clarification(
        "ai_unknown",
        "Я понял несколько разных действий сразу. Отправьте, пожалуйста, одно действие за раз.",
        state,
        chat_id,
        original_text=original_text,
    )


def handle_ai_first(
    text: str,
    planner_file: Path,
    soon_days: int,
    config: dict | None,
    state: dict | None = None,
    chat_id: int | str | None = None,
) -> str | None:
    if not config or not llm_api_key(config):
        return None
    try:
        plan = understand_action_plan(text, planner_file, config)
        reply = execute_action_plan(plan, text, planner_file, soon_days, config, state, chat_id)
        if reply is not None:
            return reply
        return request_clarification(
            "ai_unknown",
            "Не понял запрос. Сформулируйте, пожалуйста, что нужно сделать или какую выборку показать.",
            state,
            chat_id,
            original_text=text,
        )
    except Exception:
        return request_clarification(
            "ai_unknown",
            "Не получилось обработать запрос через AI. Повторите, пожалуйста, чуть позже или сформулируйте проще.",
            state,
            chat_id,
            original_text=text,
        )


def handle_pending_confirmation(text: str, state: dict, chat_id: int | str, planner_file: Path, config: dict | None = None) -> str | None:
    chat_key = str(chat_id)
    pending_delete = state.get("pending_delete", {}).get(chat_key)
    pending_move = state.get("pending_move", {}).get(chat_key)
    pending_rename = state.get("pending_rename", {}).get(chat_key)
    pending_complete = state.get("pending_complete", {}).get(chat_key)
    if not pending_delete and not pending_move and not pending_rename and not pending_complete:
        return None
    if is_stop_command(text):
        clear_pending_state(state, chat_id)
        return "Ок, ожидание сброшено."
    answer = parse_yes_no(text, config)
    if answer is True:
        if pending_delete:
            return execute_pending_delete(state, chat_id, planner_file)
        if pending_move:
            return execute_pending_move(state, chat_id, planner_file)
        if pending_rename:
            return execute_pending_rename(state, chat_id, planner_file)
        return execute_pending_complete(state, chat_id, planner_file)
    if answer is False:
        if pending_delete:
            return cancel_pending_delete(state, chat_id)
        if pending_move:
            return cancel_pending_move(state, chat_id)
        if pending_rename:
            return cancel_pending_rename(state, chat_id)
        return cancel_pending_complete(state, chat_id)
    action = "удаления" if pending_delete else "переноса" if pending_move else "переименования" if pending_rename else "отметки выполнения"
    return f"Я жду подтверждение {action}. Ответьте да или нет."


def execute_intent(
    intent: dict,
    planner_file: Path,
    soon_days: int,
    state: dict | None = None,
    chat_id: int | str | None = None,
    original_text: str | None = None,
) -> str:
    action = intent.get("action", "unknown")
    tasks = get_tasks(planner_file)

    if action == "add":
        task_name = (intent.get("task_name") or "").strip()
        due_text = (intent.get("due_date") or "").strip()
        local_adds = local_adds_from_text(original_text or "")
        if len(local_adds) == 1:
            task_name, local_due = local_adds[0]
            task = add_task(task_name, local_due, planner_file)
            return render_added_tasks([task])
        if task_name and not due_text:
            return request_add_date(task_name, state, chat_id)
        if not task_name or not due_text:
            return request_clarification(
                "add",
                intent.get("reply_hint") or "Уточните, пожалуйста, название задачи и срок.",
                state,
                chat_id,
                task_name=task_name,
                due_date=due_text,
            )
        task = add_task(task_name, parse_date(due_text), planner_file)
        return render_added_tasks([task])
    if action == "delete":
        row_numbers = [int(row) for row in intent.get("row_numbers", []) if int(row) >= 2]
        if not row_numbers and int(intent.get("row_number") or 0) >= 2:
            row_numbers = [int(intent["row_number"])]
        task_name = (intent.get("task_name") or "").strip()
        if not row_numbers and task_name:
            return request_delete_by_text(task_name, planner_file, state, chat_id, tasks)
        if not row_numbers:
            return request_clarification(
                "delete",
                intent.get("reply_hint") or "Какие номера задач удалить?",
                state,
                chat_id,
            )
        return request_delete_confirmation(row_numbers, planner_file, state, chat_id)
    if action == "move":
        row_numbers = [int(row) for row in intent.get("row_numbers", []) if int(row) >= 2]
        if not row_numbers and int(intent.get("row_number") or 0) >= 2:
            row_numbers = [int(intent["row_number"])]
        due_text = (intent.get("due_date") or "").strip()
        task_name = (intent.get("task_name") or "").strip()
        if not row_numbers and task_name and due_text:
            return request_move_by_text(task_name, parse_date(due_text), planner_file, state, chat_id, tasks)
        if not row_numbers:
            return request_clarification(
                "move",
                intent.get("reply_hint") or "Какие номера задач перенести?",
                state,
                chat_id,
                due_date=due_text,
            )
        if not due_text:
            return request_clarification(
                "move",
                intent.get("reply_hint") or "На какую дату перенести задачи: " + ", ".join(map(str, row_numbers)) + "?",
                state,
                chat_id,
                row_numbers=row_numbers,
            )
        new_due = parse_date(due_text)
        moved, missing = move_tasks(row_numbers, new_due, planner_file)
        return render_moved_tasks(moved, missing, new_due)
    if action == "rename":
        row_numbers = [int(row) for row in intent.get("row_numbers", []) if int(row) >= 2]
        if not row_numbers and int(intent.get("row_number") or 0) >= 2:
            row_numbers = [int(intent["row_number"])]
        new_name = clean_new_task_name(intent.get("new_task_name") or "")
        task_name = (intent.get("task_name") or "").strip()
        if not new_name:
            return request_clarification(
                "rename",
                intent.get("reply_hint") or "Какое новое название записать для задачи?",
                state,
                chat_id,
                row_numbers=row_numbers,
            )
        if not row_numbers and task_name:
            return request_rename_by_text(task_name, new_name, planner_file, state, chat_id, tasks)
        if not row_numbers:
            return request_clarification(
                "rename",
                intent.get("reply_hint") or "Какой номер задачи переименовать?",
                state,
                chat_id,
                new_task_name=new_name,
            )
        return request_rename_confirmation(row_numbers, new_name, planner_file, state, chat_id)
    if action == "complete":
        row_numbers = [int(row) for row in intent.get("row_numbers", []) if int(row) >= 2]
        if not row_numbers and int(intent.get("row_number") or 0) >= 2:
            row_numbers = [int(intent["row_number"])]
        task_name = (intent.get("task_name") or "").strip()
        if not row_numbers and task_name:
            return request_complete_by_text(task_name, planner_file, state, chat_id, tasks)
        if not row_numbers:
            return request_clarification(
                "complete",
                intent.get("reply_hint") or "Какие номера задач отметить выполненными?",
                state,
                chat_id,
            )
        completed_tasks, missing = complete_tasks(row_numbers, planner_file)
        return render_completed_tasks(completed_tasks, missing)
    if action == "list":
        return render_task_list(tasks)
    if action == "today":
        return render_tasks("Задачи на сегодня", select_tasks("today", tasks, soon_days))
    if action == "date":
        query_date = (intent.get("query_date") or "").strip()
        if not query_date:
            return intent.get("reply_hint") or "Не понял дату. Например: какие задачи на 12.05.2026?"
        target_date = parse_date(query_date)
        return render_tasks(f"Задачи на {format_date(target_date)}", select_tasks_for_date(tasks, target_date))
    if action == "range":
        date_from = (intent.get("date_from") or "").strip()
        date_to = (intent.get("date_to") or "").strip()
        if not date_from or not date_to:
            return intent.get("reply_hint") or "Не понял период. Например: задачи с 12.05.2026 по 15.05.2026?"
        interval = interval_between(parse_date(date_from), parse_date(date_to))
        return render_task_interval(interval, tasks)
    if action == "overdue":
        return render_tasks("Просроченные задачи", select_tasks("overdue", tasks, soon_days))
    if action == "soon":
        return render_tasks(f"Скоро, ближайшие {soon_days} дней", select_tasks("soon", tasks, soon_days))
    if action == "nearest":
        return render_tasks("Ближайшие задачи", select_tasks("nearest", tasks, soon_days))
    if action == "summary":
        return render_summary(planner_file, soon_days)
    if action == "help":
        return help_text()
    return intent.get("reply_hint") or "Не понял, что нужно сделать. Можно написать: добавить задачу, показать список или показать сводку."


def handle_text(
    text: str,
    planner_file: Path,
    soon_days: int,
    config: dict | None = None,
    state: dict | None = None,
    chat_id: int | str | None = None,
) -> str:
    low = text.lower()
    tasks = get_tasks(planner_file)
    if state is not None and chat_id is not None and is_stop_command(text):
        if has_pending_state(state, chat_id):
            clear_pending_state(state, chat_id)
            return "Ок, ожидание сброшено."
        return "Ок, сейчас я ничего не жду."
    if low.startswith(("/start", "/help", "help", "помощь")):
        return help_text()
    if state is not None and chat_id is not None:
        if has_pending_state(state, chat_id) and is_new_request(text):
            clear_pending_state(state, chat_id)
        else:
            pending_delete_reply = handle_pending_confirmation(text, state, chat_id, planner_file, config)
            if pending_delete_reply is not None:
                return pending_delete_reply
            pending_add_reply = handle_pending_add(text, state, chat_id, planner_file)
            if pending_add_reply is not None:
                return pending_add_reply
            pending_clarification_reply = handle_pending_clarification(text, state, chat_id, planner_file, soon_days, config)
            if pending_clarification_reply is not None:
                return pending_clarification_reply
    ai_first_reply = handle_ai_first(text, planner_file, soon_days, config, state, chat_id)
    if ai_first_reply is not None:
        return ai_first_reply
    parsed_rename = parse_rename_request(text)
    if parsed_rename:
        row_numbers, new_name = parsed_rename
        if row_numbers and new_name:
            return request_rename_confirmation(row_numbers, new_name, planner_file, state, chat_id)
        if row_numbers:
            return request_clarification("rename", "Какое новое название записать для задачи?", state, chat_id, row_numbers=row_numbers)
        if new_name:
            return request_rename_by_text(text, new_name, planner_file, state, chat_id, tasks)
        return request_clarification("rename", "Какую задачу переименовать и какое новое название записать?", state, chat_id)
    if low.startswith(("/add", "add", "добавить", "добавь", "+")):
        parsed_many = parse_multiple_freeform_add(text)
        if parsed_many:
            added = [add_task(name, due, planner_file) for name, due in parsed_many]
            return render_added_tasks(added)
        try:
            name, due = parse_add_command(text)
            task = add_task(name, due, planner_file)
            return render_added_tasks([task])
        except ValueError as exc:
            parsed_many = parse_multiple_freeform_add(text)
            if parsed_many:
                added = [add_task(name, due, planner_file) for name, due in parsed_many]
                return render_added_tasks(added)
            parsed_add = parse_freeform_add(text)
            if parsed_add:
                name, due = parsed_add
                if due is None:
                    return request_add_date(name, state, chat_id)
                task = add_task(name, due, planner_file)
                return render_added_tasks([task])
            if config and llm_api_key(config):
                return handle_ai_intent_or_question(text, planner_file, soon_days, config, state, chat_id)
            raise exc
    parsed_add = parse_freeform_add(text)
    parsed_many = parse_multiple_freeform_add(text)
    if parsed_many:
        added = [add_task(name, due, planner_file) for name, due in parsed_many]
        return render_added_tasks(added)
    if parsed_add:
        name, due = parsed_add
        if due is None:
            return request_add_date(name, state, chat_id)
        task = add_task(name, due, planner_file)
        return render_added_tasks([task])
    if low.startswith(("/delete", "/del", "delete", "del", "удалить", "удали")) or "удали" in low or "удалить" in low:
        return request_delete_by_text(text, planner_file, state, chat_id, tasks)
    if is_move_request(text):
        row_numbers = extract_task_numbers(text)
        new_due = extract_query_date(text)
        if not row_numbers and new_due is not None:
            return request_move_by_text(text, new_due, planner_file, state, chat_id, tasks)
        if not row_numbers or new_due is None:
            if not row_numbers and new_due is None:
                question = "Какие номера задач перенести и на какую дату?"
            elif not row_numbers:
                question = f"Какие номера задач перенести на {format_date(new_due)}?"
            else:
                question = "На какую дату перенести задачи: " + ", ".join(map(str, row_numbers)) + "?"
            return request_clarification(
                "move",
                question,
                state,
                chat_id,
                row_numbers=row_numbers,
                due_date=new_due.isoformat() if new_due else "",
            )
        moved, missing = move_tasks(row_numbers, new_due, planner_file)
        return render_moved_tasks(moved, missing, new_due)
    if is_complete_request(text):
        row_numbers = extract_task_numbers(text)
        if not row_numbers:
            return request_complete_by_text(text, planner_file, state, chat_id, tasks)
        completed_tasks, missing = complete_tasks(row_numbers, planner_file)
        return render_completed_tasks(completed_tasks, missing)
    date_interval = extract_date_interval(text)
    if date_interval is not None and not is_change_request(text):
        return render_task_interval(date_interval, tasks)
    query_date = extract_query_date(text)
    if query_date is not None and not any(word in low for word in ("добав", "напом", "запиши", "создай", "поставь")):
        return render_tasks(f"Задачи на {format_date(query_date)}", select_tasks_for_date(tasks, query_date))
    if low.startswith(("/list", "list", "список")) or "список задач" in low or "все задачи" in low:
        return render_task_list(tasks)
    if low.startswith(("/today", "сегодня")) or "сегодня" in low:
        return render_tasks("Задачи на сегодня", select_tasks("today", tasks, soon_days))
    if low.startswith(("/overdue", "просроч", "просроченные")):
        return render_tasks("Просроченные задачи", select_tasks("overdue", tasks, soon_days))
    if low.startswith(("/soon", "скоро")):
        return render_tasks(f"Скоро, ближайшие {soon_days} дней", select_tasks("soon", tasks, soon_days))
    if low.startswith(("/nearest", "ближайшие")):
        return render_tasks("Ближайшие задачи", select_tasks("nearest", tasks, soon_days))
    if low.startswith(("/summary", "сводка", "статус")):
        return render_summary(planner_file, soon_days)
    if config and llm_api_key(config):
        return handle_ai_intent_or_question(text, planner_file, soon_days, config, state, chat_id)
    return (
        "Не понял команду. Для понимания свободного текста добавьте llm_api_key в telegram_config.json.\n\n"
        + help_text()
    )


def load_state() -> dict:
    if STATE_FILE.exists():
        with STATE_FILE.open("r", encoding="utf-8") as file:
            state = json.load(file)
    else:
        state = {}
    state.setdefault("offset", 0)
    state.setdefault("last_digest_date", "")
    state.setdefault("pending_delete", {})
    state.setdefault("pending_move", {})
    state.setdefault("pending_rename", {})
    state.setdefault("pending_complete", {})
    state.setdefault("pending_add", {})
    state.setdefault("pending_clarification", {})
    return state


def save_state(state: dict) -> None:
    with STATE_FILE.open("w", encoding="utf-8") as file:
        json.dump(state, file, ensure_ascii=False, indent=2)


def maybe_send_daily_digest(config: dict, state: dict) -> None:
    digest_time = config.get("daily_digest_time", "")
    notification_chat_id = config.get("notification_chat_id", "")
    if not digest_time or not notification_chat_id:
        return
    now = datetime.now()
    today_key = now.date().isoformat()
    if state.get("last_digest_date") == today_key:
        return
    try:
        hour, minute = map(int, digest_time.split(":", 1))
    except ValueError:
        return
    if now.hour == hour and now.minute >= minute:
        planner_file = Path(config["planner_file"])
        send_message(config["bot_token"], notification_chat_id, render_summary(planner_file, int(config["soon_days"])))
        state["last_digest_date"] = today_key
        save_state(state)


def run_bot(config_path: Path = DEFAULT_CONFIG) -> None:
    config = load_config(config_path)
    token = config.get("bot_token", "")
    if not token:
        raise SystemExit("Укажите bot_token в telegram_config.json или TELEGRAM_BOT_TOKEN.")
    planner_file = Path(config["planner_file"])
    soon_days = int(config.get("soon_days", 7))
    state = load_state()
    print("Telegram task bot started. Press Ctrl+C to stop.")

    while True:
        maybe_send_daily_digest(config, state)
        response = api_request(token, "getUpdates", {
            "timeout": 55,
            "offset": int(state.get("offset", 0)),
            "allowed_updates": json.dumps(["message", "channel_post"]),
        })
        for update in response.get("result", []):
            state["offset"] = update["update_id"] + 1
            chat_id, text, message = get_update_message(update)
            has_voice = bool(message.get("voice") or message.get("audio"))
            if not text and not has_voice:
                continue
            allowed_chat_ids = config.get("allowed_chat_ids", [])
            if not allowed(chat_id, allowed_chat_ids):
                if not allowed_chat_ids and text.lower().startswith(("/start", "/help")):
                    send_message(token, chat_id, setup_lock_text(chat_id))
                continue
            try:
                if has_voice and not text:
                    transcript = transcribe_telegram_message(token, message, config)
                    command_text = normalize_voice_command(transcript)
                    pending_reply = handle_pending_confirmation(transcript, state, chat_id, planner_file, config)
                    if pending_reply is None and command_text != transcript:
                        pending_reply = handle_pending_confirmation(command_text, state, chat_id, planner_file, config)
                    reply_body = pending_reply or handle_text(command_text, planner_file, soon_days, config, state, chat_id)
                    reply = f"Распознал: {transcript}\n\n{reply_body}"
                else:
                    pending_reply = handle_pending_confirmation(text, state, chat_id, planner_file, config)
                    if pending_reply is not None:
                        reply = pending_reply
                    else:
                        reply = help_text(chat_id) if text.lower().startswith(("/start", "/help")) else handle_text(text, planner_file, soon_days, config, state, chat_id)
            except Exception as exc:
                reply = f"Не получилось обработать сообщение: {exc}"
            send_message(token, chat_id, reply)
        save_state(state)
        time.sleep(1)


def main() -> None:
    parser = argparse.ArgumentParser(description="Telegram bot for Task_Planner.xlsx")
    parser.add_argument("--config", type=Path, default=DEFAULT_CONFIG)
    parser.add_argument("--check", action="store_true", help="Show current task summary without starting Telegram")
    parser.add_argument("--add", nargs=2, metavar=("TASK", "DUE"), help="Add a task without Telegram")
    parser.add_argument("--delete", type=int, metavar="ROW", help="Delete a task row without Telegram")
    args = parser.parse_args()

    config = load_config(args.config)
    planner_file = Path(config["planner_file"])
    soon_days = int(config.get("soon_days", 7))

    if args.add:
        task_name, due_text = args.add
        task = add_task(task_name, parse_date(due_text), planner_file)
        print(f"Добавил: {task.name} — {format_date(task.due)}")
        return
    if args.delete:
        task = delete_task(args.delete, planner_file)
        print(f"Удалил: {task.name} — {format_date(task.due)}")
        return
    if args.check:
        print(render_summary(planner_file, soon_days))
        return
    run_bot(args.config)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit("\nStopped.")
