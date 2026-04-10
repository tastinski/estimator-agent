import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.drawing.image import Image as XLImage
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

    dark_fill = PatternFill("solid", fgColor="1a1a2e")
    header_fill = PatternFill("solid", fgColor="2d2d44")
    material_fill = PatternFill("solid", fgColor="fff3e0")
    labor_fill = PatternFill("solid", fgColor="e8f5e9")
    subtotal_fill = PatternFill("solid", fgColor="e8e8f0")
    total_fill = PatternFill("solid", fgColor="1a1a2e")
    discount_fill = PatternFill("solid", fgColor="fce4ec")
    white_font = Font(name="Calibri", color="FFFFFF", bold=True, size=11)
    dark_font = Font(name="Calibri", color="1a1a2e", size=11)
    bold_dark = Font(name="Calibri", color="1a1a2e", bold=True, size=11)
    total_font = Font(name="Calibri", color="FFFFFF", bold=True, size=13)
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    right_align = Alignment(horizontal="right", vertical="center")

    ws.column_dimensions["A"].width = 45
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 14

    row = 1

    logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.jpg")
    if os.path.exists(logo_path):
        img = XLImage(logo_path)
        img.width = 100
        img.height = 100
        ws.add_image(img, "A1")
        ws.row_dimensions[1].height = 75
        row = 5

    ws.merge_cells(f"A{row}:E{row}")
    ws[f"A{row}"] = "ESTIMATE"
    ws[f"A{row}"].font = Font(name="Calibri", color="FFFFFF", bold=True, size=22)
    ws[f"A{row}"].fill = dark_fill
    ws[f"A{row}"].alignment = center
    ws.row_dimensions[row].height = 45
    row += 1

    date_now = datetime.now().strftime("%B %d, %Y")
    est_num = f"EST-{datetime.now().strftime('%Y%m%d%H%M')}"
    ws.merge_cells(f"A{row}:E{row}")
    ws[f"A{row}"] = f"Date: {date_now}     Estimate #: {est_num}"
    ws[f"A{row}"].font = Font(name="Calibri", color="FFFFFF", size=10)
    ws[f"A{row}"].fill = dark_fill
    ws[f"A{row}"].alignment = center
    ws.row_dimensions[row].height = 20
    row += 2

    def get_field(pattern, text):
        match = re.search(pattern, text, re.IGNORECASE)
        return match.group(1).strip().lstrip("*").strip() if match else ""

    client = get_field(r"CLIENT:\s*(.+)", estimate_text)
    phone = get_field(r"PHONE:\s*(.+)", estimate_text)
    email = get_field(r"EMAIL:\s*(.+)", estimate_text)
    job = get_field(r"JOB:\s*(.+)", estimate_text)
    location = get_field(r"LOCATION:\s*(.+)", estimate_text)
    timeline = get_field(r"TIMELINE:\s*(.+)", estimate_text)

    ws.merge_cells(f"A{row}:E{row}")
    ws[f"A{row}"] = "PREPARED FOR"
    ws[f"A{row}"].font = bold_dark
    ws[f"A{row}"].fill = subtotal_fill
    ws[f"A{row}"].alignment = left
    ws.row_dimensions[row].height = 22
    row += 1

    for label, value in [("Client:", client), ("Phone:", phone), ("Email:", email), ("Job Type:", job), ("Location:", location)]:
        if value and value not in ["N/A", "n/a"]:
            ws.merge_cells(f"A{row}:E{row}")
            ws[f"A{row}"] = f"  {label}   {value}"
            ws[f"A{row}"].font = dark_font
            ws[f"A{row}"].alignment = left
            ws.row_dimensions[row].height = 18
            row += 1

    row += 1

    for col, header in zip(["A","B","C","D","E"], ["Description","Qty","Unit","Unit Cost","Total"]):
        ws[f"{col}{row}"] = header
        ws[f"{col}{row}"].font = white_font
        ws[f"{col}{row}"].fill = header_fill
        ws[f"{col}{row}"].alignment = center
    ws.row_dimensions[row].height = 22
    row += 1

    breakdown_match = re.search(r"BREAKDOWN:(.*?)(?:SUBTOTAL|TOTAL ESTIMATE|$)", estimate_text, re.DOTALL | re.IGNORECASE)
    material_rows = []
    labor_rows = []

    if breakdown_match:
        breakdown = breakdown_match.group(1)
        items = re.findall(r"-\s*(.+?):\s*\$?([\d,]+)", breakdown)
        for desc, cost in items:
            desc = desc.strip().lstrip("*").strip()
            cost_val = float(cost.replace(",", ""))
            is_labor = any(w in desc.lower() for w in ["labor", "labour", "hrs", "hours", "install", "work", "prep", "cleanup"])
            fill = labor_fill if is_labor else material_fill
            if is_labor:
                labor_rows.append(row)
            else:
                material_rows.append(row)

            qty = 1
            unit = "ea"
            unit_cost = cost_val
            qty_match = re.search(r"\(([\\d.]+)\s*(\w+)\s*@\s*\$?([\d.]+)\)", desc)
            if qty_match:
                qty = float(qty_match.group(1))
                unit = qty_match.group(2)
                unit_cost = float(qty_match.group(3))
                desc = re.sub(r"\s*\(.*?\)", "", desc).strip()

            ws[f"A{row}"] = f"  {desc}"
            ws[f"B{row}"] = qty
            ws[f"C{row}"] = unit
            ws[f"D{row}"] = unit_cost
            ws[f"E{row}"] = f"=B{row}*D{row}"
            for col in ["A","B","C","D","E"]:
                ws[f"{col}{row}"].font = dark_font
                ws[f"{col}{row}"].fill = fill
            ws[f"A{row}"].alignment = left
            ws[f"B{row}"].alignment = center
            ws[f"C{row}"].alignment = center
            ws[f"D{row}"].alignment = right_align
            ws[f"E{row}"].alignment = right_align
            ws[f"D{row}"].number_format = "$#,##0.00"
            ws[f"E{row}"].number_format = "$#,##0.00"
            ws.row_dimensions[row].height = 20
            row += 1

    row += 1
    subtotal_mat_row = row
    ws.merge_cells(f"A{row}:D{row}")
    ws[f"A{row}"] = "SUBTOTAL MATERIALS"
    ws[f"A{row}"].font = bold_dark
    ws[f"A{row}"].fill = material_fill
    ws[f"E{row}"] = f"=SUM(E{material_rows[0]}:E{material_rows[-1]})" if material_rows else 0
    ws[f"E{row}"].font = bold_dark
    ws[f"E{row}"].fill = material_fill
    ws[f"E{row}"].number_format = "$#,##0.00"
    ws[f"E{row}"].alignment = right_align
    ws.row_dimensions[row].height = 22
    row += 1

    subtotal_lab_row = row
    ws.merge_cells(f"A{row}:D{row}")
    ws[f"A{row}"] = "SUBTOTAL LABOR"
    ws[f"A{row}"].font = bold_dark
    ws[f"A{row}"].fill = labor_fill
    ws[f"E{row}"] = f"=SUM(E{labor_rows[0]}:E{labor_rows[-1]})" if labor_rows else 0
    ws[f"E{row}"].font = bold_dark
    ws[f"E{row}"].fill = labor_fill
    ws[f"E{row}"].number_format = "$#,##0.00"
    ws[f"E{row}"].alignment = right_align
    ws.row_dimensions[row].height = 22
    row += 1

    subtotal_row = row
    ws.merge_cells(f"A{row}:D{row}")
    ws[f"A{row}"] = "SUBTOTAL"
    ws[f"A{row}"].font = bold_dark
    ws[f"A{row}"].fill = subtotal_fill
    ws[f"E{row}"] = f"=E{subtotal_mat_row}+E{subtotal_lab_row}"
    ws[f"E{row}"].font = bold_dark
    ws[f"E{row}"].fill = subtotal_fill
    ws[f"E{row}"].number_format = "$#,##0.00"
    ws[f"E{row}"].alignment = right_align
    ws.row_dimensions[row].height = 22
    row += 1

    discount_row = row
    ws[f"A{row}"] = "DISCOUNT %"
    ws[f"A{row}"].font = bold_dark
    ws[f"A{row}"].fill = discount_fill
    ws[f"B{row}"] = 0
    ws[f"B{row}"].font = bold_dark
    ws[f"B{row}"].fill = discount_fill
    ws[f"B{row}"].number_format = "0%"
    ws[f"B{row}"].alignment = center
    ws[f"C{row}"] = "(edit %)"
    ws[f"C{row}"].font = Font(name="Calibri", color="999999", size=9, italic=True)
    ws[f"C{row}"].fill = discount_fill
    ws[f"D{row}"].fill = discount_fill
    ws[f"E{row}"] = f"=-B{row}*E{subtotal_row}"
    ws[f"E{row}"].font = bold_dark
    ws[f"E{row}"].fill = discount_fill
    ws[f"E{row}"].number_format = "$#,##0.00"
    ws[f"E{row}"].alignment = right_align
    ws.row_dimensions[row].height = 22
    row += 1

    ws.merge_cells(f"A{row}:D{row}")
    ws[f"A{row}"] = "TOTAL ESTIMATE"
    ws[f"A{row}"].font = total_font
    ws[f"A{row}"].fill = total_fill
    ws[f"E{row}"] = f"=E{subtotal_row}+E{discount_row}"
    ws[f"E{row}"].font = total_font
    ws[f"E{row}"].fill = total_fill
    ws[f"E{row}"].number_format = "$#,##0.00"
    ws[f"E{row}"].alignment = right_align
    ws.row_dimensions[row].height = 35
    row += 2

    if timeline:
        ws.merge_cells(f"A{row}:E{row}")
        ws[f"A{row}"] = f"Estimated Timeline: {timeline}"
        ws[f"A{row}"].font = dark_font
        ws.row_dimensions[row].height = 18
        row += 2

    notes_match = re.search(r"NOTES:(.*?)$", estimate_text, re.DOTALL | re.IGNORECASE)
    if notes_match:
        ws.merge_cells(f"A{row}:E{row}")
        ws[f"A{row}"] = "NOTES & CONDITIONS"
        ws[f"A{row}"].font = bold_dark
        ws[f"A{row}"].fill = subtotal_fill
        ws.row_dimensions[row].height = 20
        row += 1
        notes_text = notes_match.group(1).strip()
        notes_lines = [l.strip().lstrip("-*").strip() for l in notes_text.split("\n") if l.strip() and len(l.strip()) > 5]
        for note in notes_lines[:6]:
            ws.merge_cells(f"A{row}:E{row}")
            ws[f"A{row}"] = f"  {note}"
            ws[f"A{row}"].font = Font(name="Calibri", color="444444", size=10)
            ws[f"A{row}"].alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[row].height = 30
            row += 1

    wb.save(filepath)
    return filepath
