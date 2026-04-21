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

    # ── Page setup — Letter, fitToWidth=1 ──
    from openpyxl.worksheet.page import PageMargins
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize   = 1
    ws.page_setup.fitToPage   = True
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.5, bottom=0.5)

    ws.sheet_view.showGridLines = False
    ws.print_options.gridLines  = False

    for col_letter in "CDEFGHIJKLMNOPQRSTUVWXYZ":
        ws.column_dimensions[col_letter].hidden = True

    # ── Цвета ────────────────────────────────
    black       = "000000"
    light_gray  = "F2F2F2"
    mid_gray    = "CCCCCC"
    dark_gray   = "666666"
    header_gray = "F2F2F2"

    gray_fill   = PatternFill("solid", fgColor=light_gray)
    white_fill  = PatternFill("solid", fgColor="FFFFFF")
    header_fill = PatternFill("solid", fgColor=header_gray)

    # ── Шрифты ───────────────────────────────
    F = "Arial"
    font_num       = Font(name=F, size=12, bold=True,  color=black)
    font_title     = Font(name=F, size=20, bold=True,  color=black)
    font_date      = Font(name=F, size=10,             color=black)
    font_addr_bold = Font(name=F, size=10, bold=True,  color=black)
    font_addr      = Font(name=F, size=10,             color=black)
    font_hdr       = Font(name=F, size=10, bold=True,  color=black)
    font_item      = Font(name=F, size=10,             color=black)
    font_cost      = Font(name=F, size=10,             color=black)
    font_sub       = Font(name=F, size=11, bold=True,  color=black)
    font_total     = Font(name=F, size=12, bold=True,  color=black)
    font_note_hdr  = Font(name=F, size=10, bold=True,  color=black)
    font_note      = Font(name=F, size=10,             color=black)

    # ── Выравнивание ─────────────────────────
    left_wrap = Alignment(horizontal="left",  vertical="top",    wrap_text=True)
    left_mid  = Alignment(horizontal="left",  vertical="center", wrap_text=False)
    right_mid = Alignment(horizontal="right", vertical="center", wrap_text=False)
    right_top = Alignment(horizontal="right", vertical="top",    wrap_text=False)

    # ── Границы ──────────────────────────────
    def btm(color=mid_gray, style="thin"):
        return Border(bottom=Side(style=style, color=color))

    def full_border(color=mid_gray):
        s = Side(style="thin", color=color)
        return Border(top=s, bottom=s, left=s, right=s)

    def dbl_btm():
        return Border(bottom=Side(style="double", color=black))

    # ── Колонки — широкая A для описаний, узкая B для сумм ──
    ws.column_dimensions["A"].width = 72
    ws.column_dimensions["B"].width = 14

    # ── Парсинг полей ─────────────────────────
    def get_field(pattern, text):
        m = re.search(pattern, text, re.IGNORECASE)
        return m.group(1).strip().lstrip("*").strip() if m else ""

    def split_title_desc(raw):
        """
        Разбиваем строку на заголовок и описание.
        Убираем метки (labor)/(materials) и маркеры часов.
        """
        raw = raw.strip().lstrip("*").strip()
        raw = re.sub(r"(?i)\s*\((?:labor|labour|materials?)\)\s*", " ", raw).strip()
        raw = re.sub(r"\s*\([^)]*(?:hrs?|hours?|crew|\@\s*\$)[^)]*\)\s*", " ",
                     raw, flags=re.IGNORECASE).strip()

        for sep in [":", " — ", " – ", " - "]:
            if sep in raw:
                parts = raw.split(sep, 1)
                return parts[0].strip(), (parts[1].strip() if len(parts) > 1 else "")
        return raw, ""

    def estimate_row_height(title, desc, col_width_chars=82):
        """Оценка высоты строки под wrap_text"""
        lines = 1
        if desc:
            desc_lines = max(1, (len(desc) + col_width_chars - 1) // col_width_chars)
            lines += desc_lines + 1
        return max(22, lines * 14)

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
    material_items = []
    if breakdown_match:
        for desc, cost in re.findall(r"-\s*(.+?):\s*\$?([\d,]+)", breakdown_match.group(1)):
            raw = desc.strip()
            cost_val = float(cost.replace(",", ""))
            if re.search(r"\((?:material|materials)\)", raw, re.IGNORECASE):
                material_items.append((raw, cost_val))
            elif "labor" in raw.lower() or "labour" in raw.lower():
                labor_items.append((raw, cost_val))
            else:
                labor_items.append((raw, cost_val))

    notes_match = re.search(r"NOTES:(.*?)$", estimate_text, re.DOTALL | re.IGNORECASE)
    notes_lines = []
    if notes_match:
        notes_lines = [
            l.strip().lstrip("-*•").strip()
            for l in notes_match.group(1).strip().split("\n")
            if l.strip() and len(l.strip()) > 5
        ][:6]

    MANDATORY_NOTES = [
        "Materials and additional work (outside of the scope of work) are paid separately. All receipts will be provided by the end of the work.",
        "The estimate is valid for 30 days."
    ]
    for mn in MANDATORY_NOTES:
        if mn not in notes_lines:
            notes_lines.append(mn)

    est_num  = datetime.now().strftime("%y%m%d%H%M")
    date_now = datetime.now().strftime("%d/%m/%y")
    row = 1

    # ── ЛИНИЯ СВЕРХУ ─────────────────────────
    ws.row_dimensions[row].height = 3
    ws[f"A{row}"].border = Border(bottom=Side(style="medium", color=black))
    ws[f"B{row}"].border = Border(bottom=Side(style="medium", color=black))
    row += 1

    ws.row_dimensions[row].height = 10
    row += 1

    # ── НОМЕР ────────────────────────────────
    ws.row_dimensions[row].height = 22
    ws[f"A{row}"]           = f"#{est_num}"
    ws[f"A{row}"].font      = font_num
    ws[f"A{row}"].alignment = left_mid
    row += 1

    ws.row_dimensions[row].height = 10
    row += 1

    # ── ESTIMATE + ЛОГО ──────────────────────
    ws.row_dimensions[row].height = 60
    ws[f"A{row}"]           = "ESTIMATE"
    ws[f"A{row}"].font      = font_title
    ws[f"A{row}"].alignment = Alignment(horizontal="left", vertical="center")

    logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.jpg")
    if os.path.exists(logo_path):
        img        = XLImage(logo_path)
        img.width  = 90
        img.height = 75
        ws.add_image(img, f"B{row}")
    row += 1

    # ── DATE ─────────────────────────────────
    ws.row_dimensions[row].height = 20
    ws[f"A{row}"]           = f"DATE: {date_now}"
    ws[f"A{row}"].font      = font_date
    ws[f"A{row}"].alignment = left_mid
    row += 1

    ws.row_dimensions[row].height = 15
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
    if phone_val    and phone_val.upper()    not in ["N/A","NA",""]:
        client_lines.append((phone_val, False))
    if email_val    and email_val.upper()    not in ["N/A","NA",""]:
        client_lines.append((email_val, False))

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

    ws.row_dimensions[row].height = 20
    row += 1

    # ── ЗАГОЛОВОК ТАБЛИЦЫ ────────────────────
    ws.row_dimensions[row].height = 24
    ws[f"A{row}"]           = "Description"
    ws[f"A{row}"].font      = font_hdr
    ws[f"A{row}"].fill      = header_fill
    ws[f"A{row}"].alignment = Alignment(horizontal="left", vertical="center")
    ws[f"A{row}"].border    = full_border()
    ws[f"B{row}"]           = "Labor cost"
    ws[f"B{row}"].font      = font_hdr
    ws[f"B{row}"].fill      = header_fill
    ws[f"B{row}"].alignment = Alignment(horizontal="right", vertical="center")
    ws[f"B{row}"].border    = full_border()
    row += 1

    # ── LABOR СТРОКИ ─────────────────────────
    labor_start = row
    for i, (raw, cost_val) in enumerate(labor_items):
        title, desc = split_title_desc(raw)
        cell_text = f"- {title}\n{desc}" if desc else f"- {title}"
        fill = gray_fill if i % 2 == 0 else white_fill

        ws.row_dimensions[row].height = estimate_row_height(title, desc)
        ws[f"A{row}"]               = cell_text
        ws[f"A{row}"].font          = font_item
        ws[f"A{row}"].fill          = fill
        ws[f"A{row}"].alignment     = left_wrap
        ws[f"A{row}"].border        = btm()
        ws[f"B{row}"]               = cost_val
        ws[f"B{row}"].font          = font_cost
        ws[f"B{row}"].fill          = fill
        ws[f"B{row}"].alignment     = right_top
        ws[f"B{row}"].number_format = "$#,##0.00"
        ws[f"B{row}"].border        = btm()
        row += 1
    labor_end = row - 1

    # ── MATERIALS ────────────────────────────
    if material_items:
        ws.row_dimensions[row].height = 10
        row += 1

        ws.row_dimensions[row].height = 24
        ws[f"A{row}"]           = "Materials"
        ws[f"A{row}"].font      = font_hdr
        ws[f"A{row}"].fill      = header_fill
        ws[f"A{row}"].alignment = Alignment(horizontal="left", vertical="center")
        ws[f"A{row}"].border    = full_border()
        ws[f"B{row}"]           = "Cost"
        ws[f"B{row}"].font      = font_hdr
        ws[f"B{row}"].fill      = header_fill
        ws[f"B{row}"].alignment = Alignment(horizontal="right", vertical="center")
        ws[f"B{row}"].border    = full_border()
        row += 1

        mat_start = row
        for i, (raw, cost_val) in enumerate(material_items):
            title, desc = split_title_desc(raw)
            cell_text = f"- {title}\n{desc}" if desc else f"- {title}"
            fill = gray_fill if i % 2 == 0 else white_fill

            ws.row_dimensions[row].height = estimate_row_height(title, desc)
            ws[f"A{row}"]               = cell_text
            ws[f"A{row}"].font          = font_item
            ws[f"A{row}"].fill          = fill
            ws[f"A{row}"].alignment     = left_wrap
            ws[f"A{row}"].border        = btm()
            ws[f"B{row}"]               = cost_val
            ws[f"B{row}"].font          = font_cost
            ws[f"B{row}"].fill          = fill
            ws[f"B{row}"].alignment     = right_top
            ws[f"B{row}"].number_format = "$#,##0.00"
            ws[f"B{row}"].border        = btm()
            row += 1
        mat_end = row - 1
    else:
        mat_start = None
        mat_end = None

    ws.row_dimensions[row].height = 12
    row += 1

    # ── SUBTOTAL ─────────────────────────────
    subtotal_row = row
    ws.row_dimensions[row].height = 22
    ws[f"A{row}"]               = "Subtotal"
    ws[f"A{row}"].font          = font_sub
    ws[f"A{row}"].alignment     = right_mid
    ws[f"A{row}"].border        = btm(black, "thin")
    if material_items:
        ws[f"B{row}"] = f"=SUM(B{labor_start}:B{labor_end})+SUM(B{mat_start}:B{mat_end})"
    else:
        ws[f"B{row}"] = f"=SUM(B{labor_start}:B{labor_end})"
    ws[f"B{row}"].font          = font_sub
    ws[f"B{row}"].alignment     = right_mid
    ws[f"B{row}"].number_format = "$#,##0.00"
    ws[f"B{row}"].border        = btm(black, "thin")
    row += 1

    # ── TOTAL ────────────────────────────────
    ws.row_dimensions[row].height = 26
    ws[f"A{row}"]               = "TOTAL"
    ws[f"A{row}"].font          = font_total
    ws[f"A{row}"].alignment     = right_mid
    ws[f"A{row}"].border        = dbl_btm()
    ws[f"B{row}"]               = f"=B{subtotal_row}"
    ws[f"B{row}"].font          = font_total
    ws[f"B{row}"].alignment     = right_mid
    ws[f"B{row}"].number_format = "$#,##0.00"
    ws[f"B{row}"].border        = dbl_btm()
    row += 2

    # ── TIMELINE ─────────────────────────────
    if timeline_val:
        ws.row_dimensions[row].height = 18
        ws[f"A{row}"]           = f"Estimated Timeline: {timeline_val}"
        ws[f"A{row}"].font      = font_note
        ws[f"A{row}"].alignment = left_mid
        row += 1
        row += 1

    # ── NOTES ────────────────────────────────
    if notes_lines:
        ws.row_dimensions[row].height = 20
        ws[f"A{row}"]           = "Notes:"
        ws[f"A{row}"].font      = font_note_hdr
        ws[f"A{row}"].alignment = left_mid
        row += 1
        for note in notes_lines:
            note_lines = max(1, (len(note) + 99) // 100)
            ws.row_dimensions[row].height = max(18, note_lines * 15)
            ws[f"A{row}"]           = f"- {note}"
            ws[f"A{row}"].font      = font_note
            ws[f"A{row}"].alignment = Alignment(horizontal="left", vertical="top",
                                                wrap_text=True)
            row += 1

    ws.print_area = f"A1:B{row}"

    wb.save(filepath)
    return filepath
