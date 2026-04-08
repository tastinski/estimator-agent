import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import re
import os

def generate_estimate_excel(estimate_text: str, client_name: str, job_type: str) -> str:
    date_str = datetime.now().strftime("%Y-%m-%d")
    filename = f"estimate_{client_name.replace(' ', '_')}_{date_str}.xlsx"
    filepath = f"/tmp/{filename}"

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Estimate"

    # Стили
    dark_fill = PatternFill("solid", fgColor="1a1a2e")
    header_fill = PatternFill("solid", fgColor="2d2d44")
    row_fill = PatternFill("solid", fgColor="f8f9fa")
    total_fill = PatternFill("solid", fgColor="1a1a2e")
    white_font = Font(name="Calibri", color="FFFFFF", bold=True, size=11)
    dark_font = Font(name="Calibri", color="1a1a2e", size=11)
    total_font = Font(name="Calibri", color="FFFFFF", bold=True, size=14)
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # Ширина колонок
    ws.column_dimensions["A"].width = 60
    ws.column_dimensions["B"].width = 15

    row = 1

    # Заголовок
    ws.merge_cells(f"A{row}:B{row}")
    ws[f"A{row}"] = "ESTIMATE"
    ws[f"A{row}"].font = Font(name="Calibri", color="FFFFFF", bold=True, size=22)
    ws[f"A{row}"].fill = dark_fill
    ws[f"A{row}"].alignment = center
    ws.row_dimensions[row].height = 45
    row += 1

    # Дата и номер
    date_now = datetime.now().strftime("%B %d, %Y")
    est_num = f"EST-{datetime.now().strftime('%Y%m%d%H%M')}"
    ws.merge_cells(f"A{row}:B{row}")
    ws[f"A{row}"] = f"Date: {date_now}     Estimate #: {est_num}"
    ws[f"A{row}"].font = Font(name="Calibri", color="FFFFFF", size=10)
    ws[f"A{row}"].fill = dark_fill
    ws[f"A{row}"].alignment = center
    ws.row_dimensions[row].height = 20
    row += 2

    # Парсим данные
    def get_field(pattern, text):
        match = re.search(pattern, text, re.IGNORECASE)
        return match.group(1).strip().lstrip("*").strip() if match else ""

    client = get_field(r"CLIENT:\s*(.+)", estimate_text)
    phone = get_field(r"PHONE:\s*(.+)", estimate_text)
    email = get_field(r"EMAIL:\s*(.+)", estimate_text)
    job = get_field(r"JOB:\s*(.+)", estimate_text)
    location = get_field(r"LOCATION:\s*(.+)", estimate_text)
    timeline = get_field(r"TIMELINE:\s*(.+)", estimate_text)

    # Блок клиента
    ws.merge_cells(f"A{row}:B{row}")
    ws[f"A{row}"] = "PREPARED FOR"
    ws[f"A{row}"].font = Font(name="Calibri", color="1a1a2e", bold=True, size=12)
    ws[f"A{row}"].fill = PatternFill("solid", fgColor="e8e8f0")
    ws[f"A{row}"].alignment = left
    ws.row_dimensions[row].height = 22
    row += 1

    for label, value in [("Client:", client), ("Phone:", phone), ("Job Type:", job), ("Location:", location)]:
        if value and value != "N/A":
            ws[f"A{row}"] = f"  {label}   {value}"
            ws[f"A{row}"].font = dark_font
            ws[f"A{row}"].alignment = left
            ws.row_dimensions[row].height = 18
            row += 1

    row += 1

    # Заголовок таблицы
    ws.merge_cells(f"A{row}:B{row}")
    ws[f"A{row}"] = "SCOPE OF WORK & PRICING"
    ws[f"A{row}"].font = Font(name="Calibri", color="1a1a2e", bold=True, size=12)
    ws[f"A{row}"].fill = PatternFill("solid", fgColor="e8e8f0")
    ws[f"A{row}"].alignment = left
    ws.row_dimensions[row].height = 22
    row += 1

    # Заголовки колонок
    ws["A" + str(row)] = "Description"
    ws["B" + str(row)] = "Cost"
    for col in ["A", "B"]:
        ws[col + str(row)].font = white_font
        ws[col + str(row)].fill = header_fill
        ws[col + str(row)].alignment = center
    ws.row_dimensions[row].height = 22
    row += 1

    # Парсим breakdown
    breakdown_match = re.search(r"BREAKDOWN:(.*?)(?:TOTAL ESTIMATE|$)", estimate_text, re.DOTALL | re.IGNORECASE)
    if breakdown_match:
        breakdown = breakdown_match.group(1)
        items = re.findall(r"-\s*(.+?):\s*\$?([\d,]+)", breakdown)
        for i, (desc, cost) in enumerate(items):
            desc = desc.strip().lstrip("*").strip()
            fill = PatternFill("solid", fgColor="f8f9fa") if i % 2 == 0 else PatternFill("solid", fgColor="ffffff")
            ws[f"A{row}"] = f"  {desc}"
            ws[f"B{row}"] = f"${cost}"
            ws[f"A{row}"].font = dark_font
            ws[f"B{row}"].font = dark_font
            ws[f"A{row}"].fill = fill
            ws[f"B{row}"].fill = fill
            ws[f"A{row}"].alignment = left
            ws[f"B{row}"].alignment = center
            ws.row_dimensions[row].height = 20
            row += 1

    # Total
    total_match = re.search(r"TOTAL ESTIMATE:\s*\*{0,2}\$?([\d,]+)", estimate_text, re.IGNORECASE)
    total = total_match.group(1) if total_match else "0"

    row += 1
    ws.merge_cells(f"A{row}:B{row}")
    ws[f"A{row}"] = f"TOTAL ESTIMATE: ${total}"
    ws[f"A{row}"].font = total_font
    ws[f"A{row}"].fill = total_fill
    ws[f"A{row}"].alignment = center
    ws.row_dimensions[row].height = 35
    row += 2

    # Timeline
    if timeline:
        ws[f"A{row}"] = f"■ Estimated Timeline: {timeline}"
        ws[f"A{row}"].font = dark_font
        ws.row_dimensions[row].height = 18
        row += 2

    # Notes
    notes_match = re.search(r"NOTES:(.*?)$", estimate_text, re.DOTALL | re.IGNORECASE)
    if notes_match:
        ws.merge_cells(f"A{row}:B{row}")
        ws[f"A{row}"] = "NOTES & CONDITIONS"
        ws[f"A{row}"].font = Font(name="Calibri", color="1a1a2e", bold=True, size=11)
        ws[f"A{row}"].fill = PatternFill("solid", fgColor="e8e8f0")
        ws.row_dimensions[row].height = 20
        row += 1

        notes_text = notes_match.group(1).strip()
        notes_lines = [l.strip().lstrip("-•*").strip() for l in notes_text.split("\n") if l.strip() and not l.strip().startswith("PDF") and len(l.strip()) > 5]
        for note in notes_lines[:6]:
            ws.merge_cells(f"A{row}:B{row}")
            ws[f"A{row}"] = f"• {note}"
            ws[f"A{row}"].font = Font(name="Calibri", color="444444", size=10)
            ws[f"A{row}"].alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[row].height = 30
            row += 1

    wb.save(filepath)
    return filepath
