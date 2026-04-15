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

    # Печать — Portrait, Letter
    from openpyxl.worksheet.page import PageMargins
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize   = 1  # Letter
    ws.page_setup.fitToPage   = True
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.75, bottom=0.75)

    # ── Цвета ────────────────────────────────
    black     = "000000"
    white     = "FFFFFF"
    light_gray= "F0F0F0"
    mid_gray  = "CCCCCC"
    dark_gray = "555555"

    gray_fill    = PatternFill("solid", fgColor=light_gray)
    white_fill   = PatternFill("solid", fgColor=white)
    header_fill  = PatternFill("solid", fgColor="E8E8E8")

    # ── Шрифты ───────────────────────────────
    font_estnum  = Font(name="Calibri", size=10, bold=True,  color=black)
    font_title   = Font(name="Calibri", size=18, bold=True,  color=black)
    font_date    = Font(name="Calibri", size=10,             color=black)
    font_co_bold = Font(name="Calibri", size=10, bold=True,  color=black)
    font_co      = Font(name="Calibri", size=10,             color=black)
    font_hdr     = Font(name="Calibri", size=10, bold=True,  color=black)
    font_row     = Font(name="Calibri", size=10,             color=black)
    font_small   = Font(name="Calibri", size=9,  italic=True,color=dark_gray)
    font_sub     = Font(name="Calibri", size=10, bold=True,  color=black)
    font_total   = Font(name="Calibri", size=11, bold=True,  color=black,
                        underline="single")

    # ── Выравнивание ─────────────────────────
    left_top  = Alignment(horizontal="left",  vertical="top",    wrap_text=True)
    left_mid  = Alignment(horizontal="left",  vertical="center", wrap_text=False)
    right_mid = Alignment(horizontal="right", vertical="center")
    center_mid= Alignment(horizontal="center",vertical="center")

    # ── Границы ──────────────────────────────
    def btm(color=mid_gray, style="thin"):
        return Border(bottom=Side(style=style, color=color))
    def full(color=mid_gray):
        s = Side(style="thin", color=color)
        return Border(top=s, bottom=s, left=s, right=s)
    def dbl_btm():
        return Border(bottom=Side(style="double", color=black))
    def thick_btm():
        return Border(bottom=Side(style="medium", color=black))

    # ── Колонки ──────────────────────────────
    # A=описание (широкая), B=цена (узкая)
    ws.column_dimensions["A"].width = 72
    ws.column_dimensions["B"].width = 14

    # ── Парсинг данных ────────────────────────
    def get_field(pattern, text):
        m = re.search(pattern, text, re.IGNORECASE)
        return m.group(1).strip().lstrip("*").strip() if m else ""

    def clean_desc(desc):
        desc = re.sub(r"\s*\(?\d+[\d.]*\s*-?\s*\d*[\d.]*\s*hours?\s*@[^)]*\)?", "", desc, flags=re.IGNORECASE)
        desc = re.sub(r"\s*\(?\d+[\d.]*\s*hours?\s*[^)]*\)?",                  "", desc, flags=re.IGNORECASE)
        desc = re.sub(r"(?i)^labor\s*[-–—:]?\s*", "", desc)
        desc = re.sub(r"(?i)\s*[-–—]?\s*labor$",  "", desc)
        return desc.strip()

    client_val   = get_field(r"CLIENT:\s*(.+)",   estimate_text)
    phone_val    = get_field(r"PHONE:\s*(.+)",    estimate_text)
    email_val    = get_field(r"EMAIL:\s*(.+)",    estimate_text)
    location_val = get_field(r"LOCATION:\s*(.+)", estimate_text)
    timeline_val = get_field(r"TIMELINE:\s*(.+)", estimate_text)

    breakdown_match = re.search(
        r"BREAKDOWN:(.*?)(?:SUBTOTAL|TOTAL ESTIMATE|$)",
        estimate_text, re.DOTALL | re.IGNORECASE
    )
    labor_items    = []
    material_total = 0.0
    if breakdown_match:
        for desc, cost in re.findall(r"-\s*(.+?):\s*\$?([\d,]+)", breakdown_match.group(1)):
            desc     = desc.strip().lstrip("*").strip()
            cost_val = float(cost.replace(",", ""))
            if "labor" in desc.lower() or "labour" in desc.lower():
                labor_items.append((clean_desc(desc), cost_val))
            else:
                material_total += cost_val

    notes_match = re.search(r"NOTES:(.*?)$", estimate_text, re.DOTALL | re.IGNORECASE)
    notes_lines = []
    if notes_match:
        notes_lines = [
            l.strip().lstrip("-*•").strip()
            for l in notes_match.group(1).strip().split("\n")
            if l.strip() and len(l.strip()) > 5
        ][:8]

    est_num  = datetime.now().strftime("%y%m%d%H%M")
    date_now = datetime.now().strftime("%y/%m/%d")

    row = 1

    # ── ШАПКА ────────────────────────────────

    # Номер estimate
    ws.row_dimensions[row].height = 20
    ws[f"A{row}"]           = f"#{est_num}"
    ws[f"A{row}"].font      = font_estnum
    ws[f"A{row}"].alignment = left_mid
    row += 1

    # ESTIMATE + логотип
    ws.row_dimensions[row].height = 56
    ws[f"A{row}"]           = "ESTIMATE"
    ws[f"A{row}"].font      = font_title
    ws[f"A{row}"].alignment = Alignment(horizontal="left", vertical="center")

    logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.jpg")
    if os.path.exists(logo_path):
        img        = XLImage(logo_path)
        img.width  = 75
        img.height = 75
        ws.add_image(img, f"B{row}")
    row += 1

    # DATE
    ws.row_dimensions[row].height = 18
    ws[f"A{row}"]           = f"DATE: {date_now}"
    ws[f"A{row}"].font      = font_date
    ws[f"A{row}"].alignment = left_mid
    row += 1

    # Разделитель
    ws.row_dimensions[row].height = 6
    ws[f"A{row}"].border = btm(black, "thin")
    ws[f"B{row}"].border = btm(black, "thin")
    row += 1

    # ── ДВЕ КОЛОНКИ: подрядчик | клиент ──────
    # Делаем через вспомогательные строки
    contractor = [
        ("Qirim LLC",                            True),
        ("1105 E Orange Grove Blvd, Pasadena, CA 91104", False),
        ("adam@handyeddy.com",                   False),
        ("323 763 8686",                         False),
    ]
    client_info = []
    if client_val   and client_val.upper()   not in ["N/A","NA",""]: client_info.append((client_val,   True))
    if location_val and location_val.upper() not in ["N/A","NA",""]: client_info.append((location_val, False))
    if email_val    and email_val.upper()    not in ["N/A","NA",""]: client_info.append((email_val,    False))
    if phone_val    and phone_val.upper()    not in ["N/A","NA",""]: client_info.append((phone_val,    False))

    # Пишем в две "виртуальные" колонки внутри A и B
    # A содержит подрядчика (левая половина), B — клиента (правая)
    # Для этого используем отдельные колонки C и D — но у нас только A и B
    # Решение: делаем 3 колонки A(подрядчик широкая), пустая, B(клиент)
    # Добавим колонку C для клиента
    ws.column_dimensions["C"] = ws.column_dimensions["B"]
    ws.column_dimensions["B"].width = 2   # разделитель
    ws.column_dimensions["C"].width = 30  # клиент

    max_lines = max(len(contractor), len(client_info))
    for i in range(max_lines):
        ws.row_dimensions[row].height = 16
        if i < len(contractor):
            txt, bold = contractor[i]
            ws[f"A{row}"]           = txt
            ws[f"A{row}"].font      = Font(name="Calibri", size=10, bold=bold, color=black)
            ws[f"A{row}"].alignment = left_mid
        if i < len(client_info):
            txt, bold = client_info[i]
            ws[f"C{row}"]           = txt
            ws[f"C{row}"].font      = Font(name="Calibri", size=10, bold=bold, color=black)
            ws[f"C{row}"].alignment = left_mid
        row += 1

    ws.row_dimensions[row].height = 8
    row += 1

    # ── ЗАГОЛОВОК ТАБЛИЦЫ ────────────────────
    ws.row_dimensions[row].height = 22
    ws[f"A{row}"]           = "Description"
    ws[f"A{row}"].font      = font_hdr
    ws[f"A{row}"].fill      = header_fill
    ws[f"A{row}"].alignment = left_mid
    ws[f"A{row}"].border    = full()
    ws[f"C{row}"]           = "Labor cost"
    ws[f"C{row}"].font      = font_hdr
    ws[f"C{row}"].fill      = header_fill
    ws[f"C{row}"].alignment = right_mid
    ws[f"C{row}"].border    = full()
    row += 1

    # ── LABOR СТРОКИ ─────────────────────────
    labor_start = row
    for i, (desc, cost_val) in enumerate(labor_items):
        # Высота зависит от длины описания
        lines_needed = max(2, len(desc) // 80 + 1)
        ws.row_dimensions[row].height = max(30, lines_needed * 15)

        fill = gray_fill if i % 2 == 0 else white_fill

        ws[f"A{row}"]               = f"- {desc}"
        ws[f"A{row}"].font          = font_row
        ws[f"A{row}"].fill          = fill
        ws[f"A{row}"].alignment     = left_top
        ws[f"A{row}"].border        = btm()

        ws[f"C{row}"]               = cost_val
        ws[f"C{row}"].font          = font_row
        ws[f"C{row}"].fill          = fill
        ws[f"C{row}"].alignment     = right_mid
        ws[f"C{row}"].number_format = "$#,##0.00"
        ws[f"C{row}"].border        = btm()
        row += 1

    labor_end = row - 1

    # Materials note строка
    ws.row_dimensions[row].height = 18
    ws[f"A{row}"]               = "- Materials and additional work paid separately. All receipts provided."
    ws[f"A{row}"].font          = font_small
    ws[f"A{row}"].alignment     = left_mid
    ws[f"A{row}"].border        = btm()
    ws[f"C{row}"].border        = btm()
    row += 1

    ws.row_dimensions[row].height = 6
    row += 1

    # ── SUBTOTAL LABOR ────────────────────────
    subtotal_row = row
    ws.row_dimensions[row].height = 20
    ws[f"A{row}"]               = "Subtotal"
    ws[f"A{row}"].font          = font_sub
    ws[f"A{row}"].alignment     = left_mid
    ws[f"A{row}"].border        = thick_btm()
    ws[f"C{row}"]               = f"=SUM(C{labor_start}:C{labor_end})"
    ws[f"C{row}"].font          = font_sub
    ws[f"C{row}"].alignment     = right_mid
    ws[f"C{row}"].number_format = "$#,##0.00"
    ws[f"C{row}"].border        = thick_btm()
    row += 1

    # ── ESTIMATED MATERIALS SUBTOTAL ─────────
    mat_row = row
    ws.row_dimensions[row].height = 20
    ws[f"A{row}"]               = "Estimated Materials Subtotal"
    ws[f"A{row}"].font          = font_sub
    ws[f"A{row}"].alignment     = left_mid
    ws[f"A{row}"].border        = thick_btm()
    ws[f"C{row}"]               = material_total if material_total > 0 else 0
    ws[f"C{row}"].font          = font_sub
    ws[f"C{row}"].alignment     = right_mid
    ws[f"C{row}"].number_format = "$#,##0.00"
    ws[f"C{row}"].border        = thick_btm()
    row += 1

    # ── DISCOUNT ─────────────────────────────
    discount_row = row
    ws.row_dimensions[row].height = 20
    # Процент в A, лейбл в середине, сумма в C
    ws[f"A{row}"]               = "Discount %10"
    ws[f"A{row}"].font          = font_sub
    ws[f"A{row}"].alignment     = left_mid
    ws[f"C{row}"]               = f"=-0.10*(C{subtotal_row}+C{mat_row})"
    ws[f"C{row}"].font          = font_sub
    ws[f"C{row}"].alignment     = right_mid
    ws[f"C{row}"].number_format = "$#,##0.00"
    ws[f"C{row}"].border        = thick_btm()
    row += 1

    # ── TOTAL ────────────────────────────────
    ws.row_dimensions[row].height = 24
    ws[f"A{row}"]               = "TOTAL"
    ws[f"A{row}"].font          = font_total
    ws[f"A{row}"].alignment     = left_mid
    ws[f"A{row}"].border        = dbl_btm()
    ws[f"C{row}"]               = f"=C{subtotal_row}+C{mat_row}+C{discount_row}"
    ws[f"C{row}"].font          = font_total
    ws[f"C{row}"].alignment     = right_mid
    ws[f"C{row}"].number_format = "$#,##0.00"
    ws[f"C{row}"].border        = dbl_btm()
    row += 2

    # ── TIMELINE ─────────────────────────────
    if timeline_val:
        ws.row_dimensions[row].height = 16
        ws[f"A{row}"]           = f"Estimated Timeline: {timeline_val}"
        ws[f"A{row}"].font      = Font(name="Calibri", size=9, color=dark_gray)
        ws[f"A{row}"].alignment = left_mid
        row += 1

    row += 1

    # ── NOTES ────────────────────────────────
    if notes_lines:
        ws.row_dimensions[row].height = 16
        ws[f"A{row}"]           = "Notes:"
        ws[f"A{row}"].font      = Font(name="Calibri", size=10, bold=True,
                                       underline="single", color=black)
        ws[f"A{row}"].alignment = left_mid
        row += 1

        for note in notes_lines:
            ws.row_dimensions[row].height = 28
            ws[f"A{row}"]           = f"- {note}"
            ws[f"A{row}"].font      = Font(name="Calibri", size=9, color=dark_gray)
            ws[f"A{row}"].alignment = Alignment(horizontal="left", vertical="top",
                                                wrap_text=True)
            row += 1

    row += 1

    # ── FOOTER ───────────────────────────────
    ws.row_dimensions[row].height = 16
    ws[f"A{row}"]           = "Qirim LLC  |  adam@handyeddy.com  |  323 763 8686"
    ws[f"A{row}"].font      = Font(name="Calibri", size=9, color=dark_gray)
    ws[f"A{row}"].alignment = Alignment(horizontal="center", vertical="center")
    ws[f"A{row}"].border    = Border(top=Side(style="thin", color=mid_gray))

    # Установка области печати
    ws.print_area = f"A1:C{row}"

    wb.save(filepath)
    return filepath
