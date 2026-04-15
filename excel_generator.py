import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage
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

    # ── Цвета ────────────────────────────────
    black      = "000000"
    white      = "FFFFFF"
    light_gray = "F5F5F5"
    mid_gray   = "D0D0D0"
    dark_gray  = "404040"

    black_fill    = PatternFill("solid", fgColor=black)
    gray_fill     = PatternFill("solid", fgColor=light_gray)
    midgray_fill  = PatternFill("solid", fgColor=mid_gray)
    white_fill    = PatternFill("solid", fgColor=white)

    # ── Шрифты ───────────────────────────────
    font_title    = Font(name="Calibri", size=20, bold=True, color=black)
    font_bold     = Font(name="Calibri", size=10, bold=True, color=black)
    font_bold_wh  = Font(name="Calibri", size=10, bold=True, color=white)
    font_normal   = Font(name="Calibri", size=10, color=black)
    font_small    = Font(name="Calibri", size=9,  color=dark_gray)
    font_total    = Font(name="Calibri", size=12, bold=True, color=black, underline="single")

    # ── Выравнивание ─────────────────────────
    left   = Alignment(horizontal="left",  vertical="top", wrap_text=True)
    right  = Alignment(horizontal="right", vertical="center")
    center = Alignment(horizontal="center",vertical="center")

    # ── Границы ──────────────────────────────
    thin   = Side(style="thin",   color=mid_gray)
    thick  = Side(style="medium", color=black)

    def border(top=None, bottom=None, left=None, right=None):
        return Border(top=top, bottom=bottom, left=left, right=right)

    thin_border  = Border(top=thin, bottom=thin, left=thin, right=thin)
    thick_bottom = Border(bottom=thick)
    bottom_only  = Border(bottom=Side(style="thin", color=mid_gray))

    # ── Колонки ──────────────────────────────
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 65
    ws.column_dimensions["C"].width = 16

    # ── Парсинг данных ────────────────────────
    def get_field(pattern, text):
        match = re.search(pattern, text, re.IGNORECASE)
        return match.group(1).strip().lstrip("*").strip() if match else ""

    def clean_labor_desc(desc):
        desc = re.sub(r"\s*\(?\d+[\d.]*\s*-?\s*\d*[\d.]*\s*hours?\s*@[^)]*\)?", "", desc, flags=re.IGNORECASE)
        desc = re.sub(r"\s*\(?\d+[\d.]*\s*hours?\s*[^)]*\)?", "", desc, flags=re.IGNORECASE)
        desc = re.sub(r"(?i)^labor\s*[-–—:]?\s*", "", desc)
        desc = re.sub(r"(?i)\s*[-–—]?\s*labor$", "", desc)
        return desc.strip()

    client_val  = get_field(r"CLIENT:\s*(.+)",   estimate_text)
    phone_val   = get_field(r"PHONE:\s*(.+)",    estimate_text)
    email_val   = get_field(r"EMAIL:\s*(.+)",    estimate_text)
    job_val     = get_field(r"JOB:\s*(.+)",      estimate_text)
    location_val= get_field(r"LOCATION:\s*(.+)", estimate_text)
    timeline_val= get_field(r"TIMELINE:\s*(.+)", estimate_text)

    # Парсим breakdown
    breakdown_match = re.search(
        r"BREAKDOWN:(.*?)(?:SUBTOTAL|TOTAL ESTIMATE|$)",
        estimate_text, re.DOTALL | re.IGNORECASE
    )
    labor_items    = []
    material_total = 0.0

    if breakdown_match:
        items = re.findall(r"-\s*(.+?):\s*\$?([\d,]+)", breakdown_match.group(1))
        for desc, cost in items:
            desc     = desc.strip().lstrip("*").strip()
            cost_val = float(cost.replace(",", ""))
            if "labor" in desc.lower() or "labour" in desc.lower():
                labor_items.append((clean_labor_desc(desc), cost_val))
            else:
                material_total += cost_val

    # Парсим notes
    notes_match = re.search(r"NOTES:(.*?)$", estimate_text, re.DOTALL | re.IGNORECASE)
    notes_lines = []
    if notes_match:
        notes_lines = [
            l.strip().lstrip("-*•").strip()
            for l in notes_match.group(1).strip().split("\n")
            if l.strip() and len(l.strip()) > 5
        ][:8]

    # ── Номер estimate ────────────────────────
    est_num  = datetime.now().strftime("%y%m%d%H%M")
    date_now = datetime.now().strftime("%y/%m/%d")

    row = 1

    # ── ШАПКА ────────────────────────────────
    # Номер и дата
    ws.row_dimensions[row].height = 24
    ws.merge_cells(f"A{row}:C{row}")
    ws[f"A{row}"] = f"#{est_num}"
    ws[f"A{row}"].font      = Font(name="Calibri", size=11, bold=True, color=black)
    ws[f"A{row}"].alignment = left
    row += 1

    # ESTIMATE заголовок
    ws.row_dimensions[row].height = 36
    ws.merge_cells(f"A{row}:B{row}")
    ws[f"A{row}"]           = "ESTIMATE"
    ws[f"A{row}"].font      = font_title
    ws[f"A{row}"].alignment = Alignment(horizontal="left", vertical="center")

    # Логотип справа
    logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.jpg")
    if os.path.exists(logo_path):
        img        = XLImage(logo_path)
        img.width  = 80
        img.height = 80
        ws.add_image(img, f"C{row}")
        ws.row_dimensions[row].height = 60
    row += 1

    # DATE
    ws.row_dimensions[row].height = 18
    ws.merge_cells(f"A{row}:C{row}")
    ws[f"A{row}"]           = f"DATE: {date_now}"
    ws[f"A{row}"].font      = font_normal
    ws[f"A{row}"].alignment = left
    row += 1

    ws.row_dimensions[row].height = 8
    row += 1  # пустая строка

    # ── ДВЕ КОЛОНКИ: подрядчик и клиент ──────
    contractor_row = row
    ws.row_dimensions[row].height = 14

    # Подрядчик (левая колонка — B)
    contractor_lines = [
        "Qirim LLC",
        "1105 E Orange Grove Blvd, Pasadena, CA 91104",
        "adam@handyeddy.com",
        "323 763 8686"
    ]

    # Клиент (правая колонка — C)
    client_lines = []
    if client_val and client_val.upper() not in ["N/A", "NA", ""]:
        client_lines.append(client_val)
    if location_val and location_val.upper() not in ["N/A", "NA", ""]:
        client_lines.append(location_val)
    if email_val and email_val.upper() not in ["N/A", "NA", ""]:
        client_lines.append(email_val)
    if phone_val and phone_val.upper() not in ["N/A", "NA", ""]:
        client_lines.append(phone_val)

    max_lines = max(len(contractor_lines), len(client_lines))
    for i in range(max_lines):
        ws.row_dimensions[row].height = 16
        if i < len(contractor_lines):
            ws[f"B{row}"]           = contractor_lines[i]
            ws[f"B{row}"].font      = Font(name="Calibri", size=10,
                                           bold=(i == 0), color=black)
            ws[f"B{row}"].alignment = left
        if i < len(client_lines):
            ws[f"C{row}"]           = client_lines[i]
            ws[f"C{row}"].font      = Font(name="Calibri", size=10,
                                           bold=(i == 0), color=black)
            ws[f"C{row}"].alignment = left
        row += 1

    ws.row_dimensions[row].height = 8
    row += 1  # пустая строка

    # ── ЗАГОЛОВОК ТАБЛИЦЫ ────────────────────
    ws.row_dimensions[row].height = 22
    ws[f"B{row}"]           = "Description"
    ws[f"B{row}"].font      = font_bold
    ws[f"B{row}"].fill      = midgray_fill
    ws[f"B{row}"].alignment = left
    ws[f"B{row}"].border    = thin_border
    ws[f"C{row}"]           = "Labor cost"
    ws[f"C{row}"].font      = font_bold
    ws[f"C{row}"].fill      = midgray_fill
    ws[f"C{row}"].alignment = right
    ws[f"C{row}"].border    = thin_border
    row += 1

    # ── LABOR СТРОКИ ─────────────────────────
    labor_start = row
    for i, (desc, cost_val) in enumerate(labor_items):
        fill = gray_fill if i % 2 == 0 else white_fill
        ws.row_dimensions[row].height = 40
        ws[f"B{row}"]               = f"- {desc}"
        ws[f"B{row}"].font          = font_normal
        ws[f"B{row}"].fill          = fill
        ws[f"B{row}"].alignment     = Alignment(horizontal="left", vertical="top", wrap_text=True)
        ws[f"B{row}"].border        = bottom_only
        ws[f"C{row}"]               = cost_val
        ws[f"C{row}"].font          = font_normal
        ws[f"C{row}"].fill          = fill
        ws[f"C{row}"].alignment     = right
        ws[f"C{row}"].number_format = "$#,##0.00"
        ws[f"C{row}"].border        = bottom_only
        row += 1

    labor_end = row - 1

    # ── MATERIALS СТРОКА ─────────────────────
    ws.row_dimensions[row].height = 18
    ws[f"B{row}"]               = "- Materials and additional work paid separately. All receipts provided."
    ws[f"B{row}"].font          = Font(name="Calibri", size=9, italic=True, color=dark_gray)
    ws[f"B{row}"].alignment     = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws[f"B{row}"].border        = bottom_only
    ws[f"C{row}"]               = ""
    ws[f"C{row}"].border        = bottom_only
    row += 1

    ws.row_dimensions[row].height = 8
    row += 1  # пустая строка

    # ── SUBTOTAL ─────────────────────────────
    subtotal_row = row
    ws.row_dimensions[row].height = 20
    ws[f"B{row}"]               = "Subtotal"
    ws[f"B{row}"].font          = font_bold
    ws[f"B{row}"].alignment     = Alignment(horizontal="left", vertical="center")
    ws[f"B{row}"].border        = thick_bottom
    ws[f"C{row}"]               = f"=SUM(C{labor_start}:C{labor_end})"
    ws[f"C{row}"].font          = font_bold
    ws[f"C{row}"].alignment     = right
    ws[f"C{row}"].number_format = "$#,##0.00"
    ws[f"C{row}"].border        = thick_bottom
    row += 1

    # ── ESTIMATED MATERIALS SUBTOTAL ─────────
    mat_row = row
    ws.row_dimensions[row].height = 20
    ws[f"B{row}"]               = "Estimated Materials Subtotal"
    ws[f"B{row}"].font          = font_bold
    ws[f"B{row}"].alignment     = Alignment(horizontal="left", vertical="center")
    ws[f"B{row}"].border        = thick_bottom
    ws[f"C{row}"]               = material_total if material_total > 0 else 0
    ws[f"C{row}"].font          = font_bold
    ws[f"C{row}"].alignment     = right
    ws[f"C{row}"].number_format = "$#,##0.00"
    ws[f"C{row}"].border        = thick_bottom
    row += 1

    # ── DISCOUNT ─────────────────────────────
    discount_row = row
    ws.row_dimensions[row].height = 20

    # Процент скидки в колонке B (редактируемый)
    ws[f"B{row}"]               = "Discount %"
    ws[f"B{row}"].font          = font_bold
    ws[f"B{row}"].alignment     = Alignment(horizontal="left", vertical="center")

    # Ячейка с процентом — можно менять
    # Используем A{row} для процента
    ws[f"A{row}"]               = 0.10  # 10% по умолчанию
    ws[f"A{row}"].number_format = "0%"
    ws[f"A{row}"].font          = font_bold
    ws[f"A{row}"].alignment     = center

    ws[f"C{row}"]               = f"=-A{row}*(C{subtotal_row}+C{mat_row})"
    ws[f"C{row}"].font          = font_bold
    ws[f"C{row}"].alignment     = right
    ws[f"C{row}"].number_format = "$#,##0.00"
    ws[f"C{row}"].border        = thick_bottom
    row += 1

    # ── TOTAL ────────────────────────────────
    total_row = row
    ws.row_dimensions[row].height = 24
    ws.merge_cells(f"A{row}:B{row}")
    ws[f"A{row}"]               = "TOTAL"
    ws[f"A{row}"].font          = font_total
    ws[f"A{row}"].alignment     = Alignment(horizontal="left", vertical="center")
    ws[f"A{row}"].border        = Border(bottom=Side(style="double", color=black))
    ws[f"C{row}"]               = f"=C{subtotal_row}+C{mat_row}+C{discount_row}"
    ws[f"C{row}"].font          = font_total
    ws[f"C{row}"].alignment     = right
    ws[f"C{row}"].number_format = "$#,##0.00"
    ws[f"C{row}"].border        = Border(bottom=Side(style="double", color=black))
    row += 2

    # ── TIMELINE ─────────────────────────────
    if timeline_val:
        ws.row_dimensions[row].height = 16
        ws.merge_cells(f"A{row}:C{row}")
        ws[f"A{row}"]           = f"Estimated Timeline: {timeline_val}"
        ws[f"A{row}"].font      = font_small
        ws[f"A{row}"].alignment = left
        row += 1

    row += 1

    # ── NOTES ────────────────────────────────
    if notes_lines:
        ws.row_dimensions[row].height = 16
        ws.merge_cells(f"A{row}:C{row}")
        ws[f"A{row}"]           = "Notes:"
        ws[f"A{row}"].font      = Font(name="Calibri", size=10, bold=True,
                                       underline="single", color=black)
        ws[f"A{row}"].alignment = left
        row += 1

        for note in notes_lines:
            ws.row_dimensions[row].height = 30
            ws.merge_cells(f"A{row}:C{row}")
            ws[f"A{row}"]           = f"- {note}"
            ws[f"A{row}"].font      = font_small
            ws[f"A{row}"].alignment = Alignment(horizontal="left", vertical="top",
                                                wrap_text=True)
            row += 1

    row += 1

    # ── FOOTER ───────────────────────────────
    ws.row_dimensions[row].height = 16
    ws.merge_cells(f"A{row}:C{row}")
    ws[f"A{row}"]           = "Qirim LLC  |  adam@handyeddy.com  |  323 763 8686"
    ws[f"A{row}"].font      = Font(name="Calibri", size=9, color=dark_gray)
    ws[f"A{row}"].alignment = center
    ws[f"A{row}"].border    = Border(top=Side(style="thin", color=mid_gray))

    wb.save(filepath)
    return filepath
