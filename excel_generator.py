import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
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

    from openpyxl.worksheet.page import PageMargins
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize   = 1
    ws.page_setup.fitToPage   = True
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins = PageMargins(left=0.25, right=0.25, top=0.5, bottom=0.5)

    ws.sheet_view.showGridLines = False
    ws.print_options.gridLines  = False

    for col_letter in "CDEFGHIJKLMNOPQRSTUVWXYZ":
        ws.column_dimensions[col_letter].hidden = True

    # ── Цвета ────────────────────────────────
    black       = "000000"
    white       = "FFFFFF"
    light_gray  = "F2F2F2"
    mid_gray    = "AAAAAA"
    dark_gray   = "666666"
    header_gray = "D9D9D9"
    yellow_edit = "FFF2CC"   # жёлтый — редактируемая ячейка (дисконт)

    gray_fill   = PatternFill("solid", fgColor=light_gray)
    white_fill  = PatternFill("solid", fgColor=white)
    header_fill = PatternFill("solid", fgColor=header_gray)
    yellow_fill = PatternFill("solid", fgColor=yellow_edit)

    # ── Шрифты ───────────────────────────────
    F = "Arial"
    font_num      = Font(name=F, size=10, bold=False, color=black)
    font_title    = Font(name=F, size=16, bold=True,  color=black)
    font_date     = Font(name=F, size=10,             color=black)
    font_addr_bold= Font(name=F, size=10, bold=True,  color=black)
    font_addr     = Font(name=F, size=10,             color=black)
    font_hdr      = Font(name=F, size=10, bold=True,  color=black)
    font_row      = Font(name=F, size=10,             color=black)
    font_small    = Font(name=F, size=9,  italic=True,color=dark_gray)
    font_sub      = Font(name=F, size=10, bold=True,  color=black)
    font_total    = Font(name=F, size=11, bold=True,  color=black, underline="single")
    font_note_hdr = Font(name=F, size=10, bold=False, underline="single", color=black)
    font_note     = Font(name=F, size=9,              color=dark_gray)
    font_footer   = Font(name=F, size=9,              color=dark_gray)
    font_pct      = Font(name=F, size=10, bold=True,  color=black)

    # ── Выравнивание ─────────────────────────
    left_top  = Alignment(horizontal="left",  vertical="top",    wrap_text=True)
    left_mid  = Alignment(horizontal="left",  vertical="center", wrap_text=False)
    right_mid = Alignment(horizontal="right", vertical="center")
    center_mid= Alignment(horizontal="center",vertical="center")

    # ── Границы ──────────────────────────────
    def btm(color=mid_gray, style="thin"):
        return Border(bottom=Side(style=style, color=color))

    def full_border(color=mid_gray):
        s = Side(style="thin", color=color)
        return Border(top=s, bottom=s, left=s, right=s)

    def dbl_btm():
        return Border(bottom=Side(style="double", color=black))

    def thick_btm():
        return Border(bottom=Side(style="medium", color=black))

    # ── Колонки ──────────────────────────────
    # Делаем A чуть уже, B шире — чтобы клиент не улетал вправо
    ws.column_dimensions["A"].width = 70
    ws.column_dimensions["B"].width = 26

    # ── Парсинг ──────────────────────────────
    def get_field(pattern, text):
        m = re.search(pattern, text, re.IGNORECASE)
        return m.group(1).strip().lstrip("*").strip() if m else ""

    def clean_desc(desc):
        desc = re.sub(r"\s*\(?\d+[\d.]*\s*-?\s*\d*[\d.]*\s*hours?\s*@[^)]*\)?", "", desc, flags=re.IGNORECASE)
        desc = re.sub(r"\s*\(?\d+[\d.]*\s*hours?\s*[^)]*\)?", "", desc, flags=re.IGNORECASE)
        desc = re.sub(r"(?i)^labor\s*[-–—:]?\s*", "", desc)
        desc = re.sub(r"(?i)\s*[-–—]?\s*labor$",  "", desc)
        return desc.strip()

    client_val   = get_field(r"CLIENT:\s*(.+)",   estimate_text)
    phone_val    = get_field(r"PHONE:\s*(.+)",    estimate_text)
    email_val    = get_field(r"EMAIL:\s*(.+)",    estimate_text)
    location_val = get_field(r"LOCATION:\s*(.+)", estimate_text)
    timeline_val = get_field(r"TIMELINE:\s*(.+)", estimate_text)

    # ── Парсим breakdown — разделяем labor / materials ──
    breakdown_match = re.search(
        r"BREAKDOWN:(.*?)(?:SUBTOTAL|TOTAL ESTIMATE|$)",
        estimate_text, re.DOTALL | re.IGNORECASE
    )
    labor_items    = []
    material_items = []
    if breakdown_match:
        for desc, cost in re.findall(r"-\s*(.+?):\s*\$?([\d,]+)", breakdown_match.group(1)):
            desc     = desc.strip().lstrip("*").strip()
            cost_val = float(cost.replace(",", ""))
            if "labor" in desc.lower() or "labour" in desc.lower():
                labor_items.append((clean_desc(desc), cost_val))
            else:
                # Убираем слово "materials" из описания для чистоты
                clean = re.sub(r"(?i)\s*\(materials\)\s*", "", desc)
                clean = re.sub(r"(?i)\s*materials\s*$", "", clean)
                material_items.append((clean.strip(), cost_val))

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

    # ── ЛИНИЯ СВЕРХУ ─────────────────────────
    ws.row_dimensions[row].height = 3
    ws[f"A{row}"].border = Border(bottom=Side(style="medium", color=black))
    ws[f"B{row}"].border = Border(bottom=Side(style="medium", color=black))
    row += 1

    # ── НОМЕР ────────────────────────────────
    ws.row_dimensions[row].height = 20
    ws[f"A{row}"]           = f"#{est_num}"
    ws[f"A{row}"].font      = font_num
    ws[f"A{row}"].alignment = left_mid
    row += 1

    # ── ESTIMATE + ЛОГО ──────────────────────
    ws.row_dimensions[row].height = 70
    ws[f"A{row}"]           = "ESTIMATE"
    ws[f"A{row}"].font      = font_title
    ws[f"A{row}"].alignment = Alignment(horizontal="left", vertical="center")

    logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.jpg")
    if os.path.exists(logo_path):
        img        = XLImage(logo_path)
        img.width  = 100
        img.height = 100
        ws.add_image(img, f"B{row}")
    row += 1

    # ── DATE ─────────────────────────────────
    ws.row_dimensions[row].height = 20
    ws[f"A{row}"]           = f"DATE: {date_now}"
    ws[f"A{row}"].font      = font_date
    ws[f"A{row}"].alignment = left_mid
    row += 1

    # ── РАЗДЕЛИТЕЛЬ ──────────────────────────
    ws.row_dimensions[row].height = 3
    ws[f"A{row}"].border = btm(black, "thin")
    ws[f"B{row}"].border = btm(black, "thin")
    row += 1

    ws.row_dimensions[row].height = 10
    row += 1

    # ── АДРЕСА ───────────────────────────────
    contractor_lines = [
        ("Qirim LLC", True),
        ("1105 E Orange Grove Blvd, Pasadena, CA 91104", False),
        ("adam@handyeddy.com", False),
        ("323 763 8686", False),
    ]
    client_lines = []
    if client_val   and client_val.upper()   not in ["N/A","NA",""]:
        client_lines.append((client_val, True))
    if location_val and location_val.upper() not in ["N/A","NA",""]:
        client_lines.append((location_val, False))
    if email_val    and email_val.upper()    not in ["N/A","NA",""]:
        client_lines.append((email_val, False))
    if phone_val    and phone_val.upper()    not in ["N/A","NA",""]:
        client_lines.append((phone_val, False))

    max_addr = max(len(contractor_lines), len(client_lines), 1)
    for i in range(max_addr):
        ws.row_dimensions[row].height = 16
        if i < len(contractor_lines):
            txt, bold = contractor_lines[i]
            ws[f"A{row}"]           = txt
            ws[f"A{row}"].font      = font_addr_bold if bold else font_addr
            ws[f"A{row}"].alignment = left_mid
        if i < len(client_lines):
            txt, bold = client_lines[i]
            ws[f"B{row}"]           = txt
            ws[f"B{row}"].font      = font_addr_bold if bold else font_addr
            ws[f"B{row}"].alignment = left_mid
        row += 1

    ws.row_dimensions[row].height = 12
    row += 1

    # ── ЗАГОЛОВОК ТАБЛИЦЫ LABOR ──────────────
    ws.row_dimensions[row].height = 22
    ws[f"A{row}"]           = "Description"
    ws[f"A{row}"].font      = font_hdr
    ws[f"A{row}"].fill      = header_fill
    ws[f"A{row}"].alignment = left_mid
    ws[f"A{row}"].border    = full_border()
    ws[f"B{row}"]           = "Labor cost"
    ws[f"B{row}"].font      = font_hdr
    ws[f"B{row}"].fill      = header_fill
    ws[f"B{row}"].alignment = right_mid
    ws[f"B{row}"].border    = full_border()
    row += 1

    # ── LABOR СТРОКИ ─────────────────────────
    labor_start = row
    for i, (desc, cost_val) in enumerate(labor_items):
        lines_needed = max(1, len(desc) // 70 + 1)
        ws.row_dimensions[row].height = max(18, lines_needed * 14)
        fill = gray_fill if i % 2 == 0 else white_fill

        ws[f"A{row}"]               = f"- {desc}"
        ws[f"A{row}"].font          = font_row
        ws[f"A{row}"].fill          = fill
        ws[f"A{row}"].alignment     = left_top
        ws[f"A{row}"].border        = btm()
        ws[f"B{row}"]               = cost_val
        ws[f"B{row}"].font          = font_row
        ws[f"B{row}"].fill          = fill
        ws[f"B{row}"].alignment     = right_mid
        ws[f"B{row}"].number_format = "$#,##0.00"
        ws[f"B{row}"].border        = btm()
        row += 1

    labor_end = row - 1

    ws.row_dimensions[row].height = 6
    row += 1

    # ── ЗАГОЛОВОК MATERIALS ──────────────────
    if material_items:
        ws.row_dimensions[row].height = 22
        ws[f"A{row}"]           = "Materials"
        ws[f"A{row}"].font      = font_hdr
        ws[f"A{row}"].fill      = header_fill
        ws[f"A{row}"].alignment = left_mid
        ws[f"A{row}"].border    = full_border()
        ws[f"B{row}"]           = "Cost"
        ws[f"B{row}"].font      = font_hdr
        ws[f"B{row}"].fill      = header_fill
        ws[f"B{row}"].alignment = right_mid
        ws[f"B{row}"].border    = full_border()
        row += 1

        # ── MATERIALS СТРОКИ ─────────────────
        mat_start = row
        for i, (desc, cost_val) in enumerate(material_items):
            lines_needed = max(1, len(desc) // 70 + 1)
            ws.row_dimensions[row].height = max(18, lines_needed * 14)
            fill = gray_fill if i % 2 == 0 else white_fill

            ws[f"A{row}"]               = f"- {desc}"
            ws[f"A{row}"].font          = font_row
            ws[f"A{row}"].fill          = fill
            ws[f"A{row}"].alignment     = left_top
            ws[f"A{row}"].border        = btm()
            ws[f"B{row}"]               = cost_val
            ws[f"B{row}"].font          = font_row
            ws[f"B{row}"].fill          = fill
            ws[f"B{row}"].alignment     = right_mid
            ws[f"B{row}"].number_format = "$#,##0.00"
            ws[f"B{row}"].border        = btm()
            row += 1
        mat_end = row - 1
    else:
        # Если materials отсутствуют — пустой диапазон для формулы
        mat_start = row
        mat_end   = row - 1  # пустой диапазон, SUM вернёт 0

    ws.row_dimensions[row].height = 6
    row += 1

    # ── SUBTOTAL LABOR ───────────────────────
    subtotal_row = row
    ws.row_dimensions[row].height = 20
    ws[f"A{row}"]               = "Subtotal Labor"
    ws[f"A{row}"].font          = font_sub
    ws[f"A{row}"].alignment     = left_mid
    ws[f"A{row}"].border        = thick_btm()
    ws[f"B{row}"]               = f"=SUM(B{labor_start}:B{labor_end})"
    ws[f"B{row}"].font          = font_sub
    ws[f"B{row}"].alignment     = right_mid
    ws[f"B{row}"].number_format = "$#,##0.00"
    ws[f"B{row}"].border        = thick_btm()
    row += 1

    # ── SUBTOTAL MATERIALS (формула!) ───────
    mat_row = row
    ws.row_dimensions[row].height = 20
    ws[f"A{row}"]               = "Subtotal Materials"
    ws[f"A{row}"].font          = font_sub
    ws[f"A{row}"].alignment     = left_mid
    ws[f"A{row}"].border        = thick_btm()
    if material_items:
        ws[f"B{row}"] = f"=SUM(B{mat_start}:B{mat_end})"
    else:
        ws[f"B{row}"] = 0
    ws[f"B{row}"].font          = font_sub
    ws[f"B{row}"].alignment     = right_mid
    ws[f"B{row}"].number_format = "$#,##0.00"
    ws[f"B{row}"].border        = thick_btm()
    row += 1

    # ── DISCOUNT — редактируемый процент ────
    # Ячейка A с процентом отдельно, в B формула.
    # Adam меняет процент в A — пересчёт автоматически.
    discount_row = row
    ws.row_dimensions[row].height = 20

    # В колонке A — "Discount" + жёлтая ячейка с процентом сразу после
    # Делаем через единый текст с процентом, но процент в именованной ячейке
    # Для простоты ставим всё в A и процент в отдельную ячейку справа через формулу
    ws[f"A{row}"]               = "Discount (edit % →)"
    ws[f"A{row}"].font          = font_sub
    ws[f"A{row}"].alignment     = left_mid

    # Процент вписываем в A но сразу после слова — не получится. Используем helper:
    # Вместо этого — храним процент в специальной ячейке в конце документа,
    # на неё ссылаемся через formula. Но проще: процент в ячейке A, а текст перед ним.
    # Решение: процент в скрытой ячейке C{discount_row}, в B — формула
    ws[f"C{discount_row}"] = 0.10  # 10% по умолчанию — редактируемое значение
    ws[f"C{discount_row}"].number_format = "0%"
    ws[f"C{discount_row}"].font          = font_pct
    ws[f"C{discount_row}"].fill          = yellow_fill
    ws[f"C{discount_row}"].alignment     = center_mid

    # Разкрываем колонку C чтобы Adam видел процент
    ws.column_dimensions["C"].hidden = False
    ws.column_dimensions["C"].width  = 10

    # Формула в A: "Discount [10%]:"
    ws[f"A{row}"] = f'=CONCATENATE("Discount (",TEXT(C{discount_row},"0%")," — edit yellow cell →)")'

    # Сумма discount в B
    ws[f"B{row}"]               = f"=-C{discount_row}*(B{subtotal_row}+B{mat_row})"
    ws[f"B{row}"].font          = font_sub
    ws[f"B{row}"].alignment     = right_mid
    ws[f"B{row}"].number_format = "$#,##0.00"
    ws[f"B{row}"].border        = thick_btm()
    row += 1

    # ── TOTAL ────────────────────────────────
    ws.row_dimensions[row].height = 24
    ws[f"A{row}"]               = "TOTAL"
    ws[f"A{row}"].font          = font_total
    ws[f"A{row}"].alignment     = left_mid
    ws[f"A{row}"].border        = dbl_btm()
    ws[f"B{row}"]               = f"=B{subtotal_row}+B{mat_row}+B{discount_row}"
    ws[f"B{row}"].font          = font_total
    ws[f"B{row}"].alignment     = right_mid
    ws[f"B{row}"].number_format = "$#,##0.00"
    ws[f"B{row}"].border        = dbl_btm()
    row += 2

    # ── TIMELINE ─────────────────────────────
    if timeline_val:
        ws.row_dimensions[row].height = 15
        ws[f"A{row}"]           = f"Estimated Timeline: {timeline_val}"
        ws[f"A{row}"].font      = font_note
        ws[f"A{row}"].alignment = left_mid
        row += 1

    row += 1

    # ── NOTES ────────────────────────────────
    if notes_lines:
        ws.row_dimensions[row].height = 15
        ws[f"A{row}"]           = "Notes:"
        ws[f"A{row}"].font      = font_note_hdr
        ws[f"A{row}"].alignment = left_mid
        row += 1
        for note in notes_lines:
            ws.row_dimensions[row].height = 25
            ws[f"A{row}"]           = f"- {note}"
            ws[f"A{row}"].font      = font_note
            ws[f"A{row}"].alignment = Alignment(horizontal="left", vertical="top",
                                                wrap_text=True)
            row += 1

    row += 1

    # ── FOOTER ───────────────────────────────
    ws.row_dimensions[row].height = 3
    ws[f"A{row}"].border = Border(bottom=Side(style="thin", color=black))
    ws[f"B{row}"].border = Border(bottom=Side(style="thin", color=black))
    row += 1

    ws.row_dimensions[row].height = 15
    ws.merge_cells(f"A{row}:B{row}")
    ws[f"A{row}"]           = "Qirim LLC  |  adam@handyeddy.com  |  323 763 8686"
    ws[f"A{row}"].font      = font_footer
    ws[f"A{row}"].alignment = center_mid

    # ── Print area — включаем C чтобы процент дисконта печатался ──
    ws.print_area = f"A1:C{row}"

    wb.save(filepath)
    return filepath
