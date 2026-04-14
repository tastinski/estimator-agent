```python
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

    orange     = "F97316"
    gray       = "9CA3AF"
    light_gray = "F3F4F6"
    white      = "FFFFFF"
    dark_text  = "1F2937"

    orange_fill = PatternFill("solid", fgColor=orange)
    gray_fill   = PatternFill("solid", fgColor=gray)
    light_fill  = PatternFill("solid", fgColor=light_gray)
    white_fill  = PatternFill("solid", fgColor=white)

    white_bold  = Font(name="Calibri", color=white, bold=True, size=11)
    white_large = Font(name="Calibri", color=white, bold=True, size=24)
    dark_font   = Font(name="Calibri", color=dark_text, size=11)
    dark_bold   = Font(name="Calibri", color=dark_text, bold=True, size=11)

    center      = Alignment(horizontal="center", vertical="center")
    left        = Alignment(horizontal="left", vertical="center", wrap_text=True)
    right_align = Alignment(horizontal="right", vertical="center")

    ws.column_dimensions["A"].width = 75
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 16

    row = 1

    # Логотип
    logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.jpg")
    if os.path.exists(logo_path):
        img        = XLImage(logo_path)
        img.width  = 90
        img.height = 90
        ws.add_image(img, "D1")
        ws.row_dimensions[1].height = 70
        ws.row_dimensions[2].height = 20
        ws.row_dimensions[3].height = 20
        row = 4

    # Шапка — оранжевая
    ws.merge_cells(f"A{row}:E{row}")
    ws[f"A{row}"]           = "HANDY EDDY HANDYMAN"
    ws[f"A{row}"].font      = white_large
    ws[f"A{row}"].fill      = orange_fill
    ws[f"A{row}"].alignment = center
    ws.row_dimensions[row].height = 50
    row += 1

    ws.merge_cells(f"A{row}:E{row}")
    ws[f"A{row}"]           = "ESTIMATE"
    ws[f"A{row}"].font      = Font(name="Calibri", color=white, bold=True, size=14)
    ws[f"A{row}"].fill      = gray_fill
    ws[f"A{row}"].alignment = center
    ws.row_dimensions[row].height = 25
    row += 1

    date_now = datetime.now().strftime("%B %d, %Y")
    est_num  = f"EST-{datetime.now().strftime('%Y%m%d%H%M')}"
    ws.merge_cells(f"A{row}:E{row}")
    ws[f"A{row}"]           = f"Date: {date_now}     Estimate #: {est_num}"
    ws[f"A{row}"].font      = Font(name="Calibri", color=dark_text, size=10)
    ws[f"A{row}"].fill      = light_fill
    ws[f"A{row}"].alignment = center
    ws.row_dimensions[row].height = 20
    row += 2

    def get_field(pattern, text):
        match = re.search(pattern, text, re.IGNORECASE)
        return match.group(1).strip().lstrip("*").strip() if match else ""

    client   = get_field(r"CLIENT:\s*(.+)", estimate_text)
    phone    = get_field(r"PHONE:\s*(.+)", estimate_text)
    email    = get_field(r"EMAIL:\s*(.+)", estimate_text)
    job      = get_field(r"JOB:\s*(.+)", estimate_text)
    location = get_field(r"LOCATION:\s*(.+)", estimate_text)
    timeline = get_field(r"TIMELINE:\s*(.+)", estimate_text)

    # Блок клиента — оранжевый заголовок
    ws.merge_cells(f"A{row}:E{row}")
    ws[f"A{row}"]           = "PREPARED FOR"
    ws[f"A{row}"].font      = white_bold
    ws[f"A{row}"].fill      = orange_fill
    ws[f"A{row}"].alignment = left
    ws.row_dimensions[row].height = 22
    row += 1

    for label, value in [
        ("Client:", client), ("Phone:", phone), ("Email:", email),
        ("Job Type:", job), ("Location:", location)
    ]:
        if value and value.upper() not in ["N/A", "NA", "[PENDING]", "PENDING", ""]:
            ws.merge_cells(f"A{row}:E{row}")
            ws[f"A{row}"]           = f"  {label}   {value}"
            ws[f"A{row}"].font      = dark_font
            ws[f"A{row}"].fill      = white_fill
            ws[f"A{row}"].alignment = left
            ws.row_dimensions[row].height = 18
            row += 1

    row += 1

    # Заголовки колонок — серые
    for col, header in zip(["A","B","C","D","E"],
                           ["Description", "Qty", "Unit", "Unit Cost", "Total"]):
        ws[f"{col}{row}"]           = header
        ws[f"{col}{row}"].font      = white_bold
        ws[f"{col}{row}"].fill      = gray_fill
        ws[f"{col}{row}"].alignment = center
    ws.row_dimensions[row].height = 22
    row += 1

    # Парсим breakdown
    breakdown_match = re.search(
        r"BREAKDOWN:(.*?)(?:SUBTOTAL|TOTAL ESTIMATE|$)",
        estimate_text, re.DOTALL | re.IGNORECASE
    )

    labor_rows     = []
    material_total = 0.0

    if breakdown_match:
        breakdown = breakdown_match.group(1)
        items     = re.findall(r"-\s*(.+?):\s*\$?([\d,]+)", breakdown)

        for desc, cost in items:
            desc     = desc.strip().lstrip("*").strip()
            cost_val = float(cost.replace(",", ""))
            is_labor = any(w in desc.lower() for w in [
                "labor", "labour", "hrs", "hours", "install", "work",
                "prep", "cleanup", "washing", "scraping", "sanding",
                "painting", "priming", "removal", "replacement",
                "trim", "casing", "touch-up", "touchup"
            ])

            if is_labor:
                # Labor — каждая строка отдельно, белый/светло-серый чередование
                fill = light_fill if len(labor_rows) % 2 == 0 else white_fill
                labor_rows.append(row)

                qty       = 1
                unit      = "ea"
                unit_cost = cost_val
                qty_match = re.search(r"\(([0-9.]+)\s*(\w+)\s*@\s*\$?([0-9.]+)\)", desc)
                if qty_match:
                    qty       = float(qty_match.group(1))
                    unit      = qty_match.group(2)
                    unit_cost = float(qty_match.group(3))
                    desc      = re.sub(r"\s*\(.*?\)", "", desc).strip()

                ws[f"A{row}"]               = f"  {desc}"
                ws[f"B{row}"]               = qty
                ws[f"C{row}"]               = unit
                ws[f"D{row}"]               = unit_cost
                ws[f"E{row}"]               = f"=B{row}*D{row}"
                ws[f"A{row}"].alignment     = left
                ws[f"B{row}"].alignment     = center
                ws[f"C{row}"].alignment     = center
                ws[f"D{row}"].alignment     = right_align
                ws[f"E{row}"].alignment     = right_align
                ws[f"D{row}"].number_format = "$#,##0.00"
                ws[f"E{row}"].number_format = "$#,##0.00"
                for col in ["A","B","C","D","E"]:
                    ws[f"{col}{row}"].font = dark_font
                    ws[f"{col}{row}"].fill = fill
                ws.row_dimensions[row].height = 20
                row += 1
            else:
                # Материалы — только суммируем
                material_total += cost_val

    # Одна строка Materials с итогом
    mat_row = row
    ws.merge_cells(f"A{row}:D{row}")
    ws[f"A{row}"]               = "  Materials (total)"
    ws[f"A{row}"].font          = dark_font
    ws[f"A{row}"].fill          = white_fill
    ws[f"A{row}"].alignment     = left
    ws[f"E{row}"]               = material_total
    ws[f"E{row}"].font          = dark_font
    ws[f"E{row}"].fill          = white_fill
    ws[f"E{row}"].number_format = "$#,##0.00"
    ws[f"E{row}"].alignment     = right_align
    ws.row_dimensions[row].height = 20
    row += 2

    # Subtotal Materials
    subtotal_mat_row = row
    ws.merge_cells(f"A{row}:D{row}")
    ws[f"A{row}"]               = "SUBTOTAL MATERIALS"
    ws[f"A{row}"].font          = dark_bold
    ws[f"A{row}"].fill          = light_fill
    ws[f"E{row}"]               = material_total
    ws[f"E{row}"].font          = dark_bold
    ws[f"E{row}"].fill          = light_fill
    ws[f"E{row}"].number_format = "$#,##0.00"
    ws[f"E{row}"].alignment     = right_align
    ws.row_dimensions[row].height = 22
    row += 1

    # Subtotal Labor — серый, без оранжевого
    subtotal_lab_row = row
    ws.merge_cells(f"A{row}:D{row}")
    ws[f"A{row}"]               = "SUBTOTAL LABOR"
    ws[f"A{row}"].font          = dark_bold
    ws[f"A{row}"].fill          = light_fill
    ws[f"E{row}"]               = f"=SUM(E{labor_rows[0]}:E{labor_rows[-1]})" if labor_rows else 0
    ws[f"E{row}"].font          = dark_bold
    ws[f"E{row}"].fill          = light_fill
    ws[f"E{row}"].number_format = "$#,##0.00"
    ws[f"E{row}"].alignment     = right_align
    ws.row_dimensions[row].height = 22
    row += 1

    # Subtotal
    subtotal_row = row
    ws.merge_cells(f"A{row}:D{row}")
    ws[f"A{row}"]               = "SUBTOTAL"
    ws[f"A{row}"].font          = dark_bold
    ws[f"A{row}"].fill          = light_fill
    ws[f"E{row}"]               = f"=E{subtotal_mat_row}+E{subtotal_lab_row}"
    ws[f"E{row}"].font          = dark_bold
    ws[f"E{row}"].fill          = light_fill
    ws[f"E{row}"].number_format = "$#,##0.00"
    ws[f"E{row}"].alignment     = right_align
    ws.row_dimensions[row].height = 22
    row += 1

    # Discount — без оранжевого шрифта
    discount_row = row
    ws[f"A{row}"]               = "DISCOUNT %"
    ws[f"A{row}"].font          = dark_bold
    ws[f"A{row}"].fill          = white_fill
    ws[f"B{row}"]               = 0
    ws[f"B{row}"].font          = dark_bold
    ws[f"B{row}"].fill          = white_fill
    ws[f"B{row}"].number_format = "0%"
    ws[f"B{row}"].alignment     = center
    ws[f"C{row}"]               = "(edit %)"
    ws[f"C{row}"].font          = Font(name="Calibri", color=gray, size=9, italic=True)
    ws[f"C{row}"].fill          = white_fill
    ws[f"D{row}"].fill          = white_fill
    ws[f"E{row}"]               = f"=-B{row}*E{subtotal_row}"
    ws[f"E{row}"].font          = dark_bold
    ws[f"E{row}"].fill          = white_fill
    ws[f"E{row}"].number_format = "$#,##0.00"
    ws[f"E{row}"].alignment     = right_align
    ws.row_dimensions[row].height = 22
    row += 1

    # Total — оранжевый только здесь
    ws.merge_cells(f"A{row}:D{row}")
    ws[f"A{row}"]               = "TOTAL ESTIMATE"
    ws[f"A{row}"].font          = Font(name="Calibri", color=white, bold=True, size=14)
    ws[f"A{row}"].fill          = orange_fill
    ws[f"A{row}"].alignment     = center
    ws[f"E{row}"]               = f"=E{subtotal_row}+E{discount_row}"
    ws[f"E{row}"].font          = Font(name="Calibri", color=white, bold=True, size=14)
    ws[f"E{row}"].fill          = orange_fill
    ws[f"E{row}"].number_format = "$#,##0.00"
    ws[f"E{row}"].alignment     = right_align
    ws.row_dimensions[row].height = 35
    row += 2

    if timeline:
        ws.merge_cells(f"A{row}:E{row}")
        ws[f"A{row}"]      = f"Estimated Timeline: {timeline}"
        ws[f"A{row}"].font = dark_font
        ws.row_dimensions[row].height = 18
        row += 2

    # Notes
    notes_match = re.search(r"NOTES:(.*?)$", estimate_text, re.DOTALL | re.IGNORECASE)
    if notes_match:
        ws.merge_cells(f"A{row}:E{row}")
        ws[f"A{row}"]      = "NOTES & CONDITIONS"
        ws[f"A{row}"].font = white_bold
        ws[f"A{row}"].fill = gray_fill
        ws.row_dimensions[row].height = 20
        row += 1
        notes_text  = notes_match.group(1).strip()
        notes_lines = [
            l.strip().lstrip("-*").strip()
            for l in notes_text.split("\n")
            if l.strip() and len(l.strip()) > 5
        ]
        for note in notes_lines[:6]:
            ws.merge_cells(f"A{row}:E{row}")
            ws[f"A{row}"]           = f"  • {note}"
            ws[f"A{row}"].font      = Font(name="Calibri", color=dark_text, size=10)
            ws[f"A{row}"].fill      = white_fill
            ws[f"A{row}"].alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[row].height = 30
            row += 1

    # Footer — серый
    row += 1
    ws.merge_cells(f"A{row}:E{row}")
    ws[f"A{row}"]           = "Handy Eddy Handyman  |  adam@handyeddy.com  |  (323) 763-8686"
    ws[f"A{row}"].font      = Font(name="Calibri", color=white, size=10)
    ws[f"A{row}"].fill      = gray_fill
    ws[f"A{row}"].alignment = center
    ws.row_dimensions[row].height = 20

    wb.save(filepath)
    return filepath
```

Коммить.
