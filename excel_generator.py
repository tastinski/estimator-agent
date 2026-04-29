import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage
from datetime import datetime
import re
import os


def generate_estimate_excel(estimate_text: str, client_name: str, job_type: str, card_number: str = None) -> str:
    """
    Генерим Excel эстимейт.
    card_number — номер из Monday карточки (text5__1) или папки Dropbox.
    Если передан — используется в имени файла и в номере на бланке.
    Если нет — fallback на datetime (но это не должно происходить в продакшене).
    """
    # Имя файла = тот же номер что у папки в Dropbox
    if card_number:
        safe_name = client_name.strip() or "Client"
        filename = f"{card_number} {safe_name} estimate.xlsx"
    else:
        date_str = datetime.now().strftime("%Y-%m-%d")
        filename = f"estimate_{client_name.replace(' ', '_')}_{date_str}.xlsx"
    filepath = f"/tmp/{filename}"

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Estimate"

    from openpyxl.worksheet.page import PageMargins
    ws.page_setup.orientation  = "portrait"
    ws.page_setup.paperSize    = 1
    ws.page_setup.fitToPage    = True
    ws.page_setup.fitToWidth   = 1
    ws.page_setup.fitToHeight  = 0
    ws.page_margins = PageMargins(left=0.25, right=0.25, top=0.4, bottom=0.4)
    ws.sheet_view.showGridLines = False
    ws.print_options.gridLines  = False

    # Скрываем все колонки кроме A и B (C используется для редактируемого discount %)
    for col in "CDEFGHIJKLMNOPQRSTUVWXYZ":
        ws.column_dimensions[col].hidden = True

    # ── Цвета ────────────────────────────────
    BLACK      = "000000"
    DARK_GRAY  = "666666"
    LIGHT_GRAY = "F2F2F2"
    MID_GRAY   = "CCCCCC"

    gray_fill   = PatternFill("solid", fgColor=LIGHT_GRAY)
    white_fill  = PatternFill("solid", fgColor="FFFFFF")
    header_fill = PatternFill("solid", fgColor="D9D9D9")

    # ── Шрифты Arial ─────────────────────────
    F = "Arial"
    def fnt(size=10, bold=False, color=BLACK, underline=False, italic=False):
        return Font(name=F, size=size, bold=bold, color=color,
                    underline="single" if underline else None, italic=italic)

    # ── Выравнивание ─────────────────────────
    L  = Alignment(horizontal="left",   vertical="center", wrap_text=False)
    LW = Alignment(horizontal="left",   vertical="top",    wrap_text=True)
    R  = Alignment(horizontal="right",  vertical="center", wrap_text=False)
    RT = Alignment(horizontal="right",  vertical="top",    wrap_text=False)
    C  = Alignment(horizontal="center", vertical="center", wrap_text=False)

    # ── Границы ──────────────────────────────
    def btm(color=MID_GRAY, style="thin"):
        return Border(bottom=Side(style=style, color=color))
    def full(color=MID_GRAY):
        s = Side(style="thin", color=color)
        return Border(top=s, bottom=s, left=s, right=s)
    def dbl():
        return Border(bottom=Side(style="double", color=BLACK))
    def med():
        return Border(bottom=Side(style="medium", color=BLACK))

    # ── Колонки ──────────────────────────────
    ws.column_dimensions["A"].width = 68
    ws.column_dimensions["B"].width = 20

    # ── Парсинг ──────────────────────────────
    def get_field(pattern, text):
        m = re.search(pattern, text, re.IGNORECASE)
        return m.group(1).strip().lstrip("*").strip() if m else ""

    client_val   = get_field(r"CLIENT:\s*(.+)",   estimate_text)
    phone_val    = get_field(r"PHONE:\s*(.+)",    estimate_text)
    email_val    = get_field(r"EMAIL:\s*(.+)",    estimate_text)
    location_val = get_field(r"LOCATION:\s*(.+)", estimate_text)
    timeline_val = get_field(r"TIMELINE:\s*(.+)", estimate_text)

    # Парсим breakdown — разделяем labor и materials
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
            # Чистим лейблы (labor), (materials) и описания часов
            clean = re.sub(r"\s*\([^)]*(?:hrs?|hours?|crew|@\s*\$)[^)]*\)", "", desc, flags=re.IGNORECASE)
            clean = re.sub(r"\s*\((?:labor|labour)\)\s*", " ", clean, flags=re.IGNORECASE).strip()
            clean = re.sub(r"\s*[-–—]?\s*(?:labor|labour)$", "", clean, flags=re.IGNORECASE).strip()

            if re.search(r"\(materials?\)", desc, re.IGNORECASE):
                mat_clean = re.sub(r"\s*\(materials?\)\s*", " ", clean, flags=re.IGNORECASE).strip()
                material_items.append((mat_clean, cost_val))
            elif re.search(r"\bmaterials?\b", desc, re.IGNORECASE) and not re.search(r"\blabor\b", desc, re.IGNORECASE):
                material_items.append((clean, cost_val))
            else:
                labor_items.append((clean, cost_val))

    # Notes — только из текста агента, БЕЗ автоматических mandatory нотов
    notes_match = re.search(r"NOTES:(.*?)$", estimate_text, re.DOTALL | re.IGNORECASE)
    notes_lines = []
    if notes_match:
        notes_lines = [
            l.strip().lstrip("-*•").strip()
            for l in notes_match.group(1).strip().split("\n")
            if l.strip() and len(l.strip()) > 5
        ][:8]

    # Номер на бланке = card_number из Monday/Dropbox
    est_num  = card_number if card_number else datetime.now().strftime("%y%m%d%H%M")
    date_now = datetime.now().strftime("%d/%m/%y")
    row = 1

    # ════════════════════════════════════════
    # ШАПКА
    # ════════════════════════════════════════

    # Верхняя линия
    ws.row_dimensions[row].height = 3
    for col in ["A", "B"]:
        ws[f"{col}{row}"].border = Border(bottom=Side(style="medium", color=BLACK))
    row += 1

    ws.row_dimensions[row].height = 8
    row += 1

    # Номер
    ws.row_dimensions[row].height = 20
    ws[f"A{row}"]           = f"#{est_num}"
    ws[f"A{row}"].font      = fnt(10, bold=True)
    ws[f"A{row}"].alignment = L
    row += 1

    ws.row_dimensions[row].height = 8
    row += 1

    # ESTIMATE + лого
    ws.row_dimensions[row].height = 65
    ws[f"A{row}"]           = "ESTIMATE"
    ws[f"A{row}"].font      = fnt(18, bold=True)
    ws[f"A{row}"].alignment = Alignment(horizontal="left", vertical="center")

    logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.jpg")
    if os.path.exists(logo_path):
        img = XLImage(logo_path)
        img.width  = 90
        img.height = 90
        ws.add_image(img, f"B{row}")
    row += 1

    # DATE
    ws.row_dimensions[row].height = 20
    ws[f"A{row}"]           = f"DATE: {date_now}"
    ws[f"A{row}"].font      = fnt(10)
    ws[f"A{row}"].alignment = L
    row += 1

    ws.row_dimensions[row].height = 10
    row += 1

    # Разделитель
    ws.row_dimensions[row].height = 3
    for col in ["A", "B"]:
        ws[f"{col}{row}"].border = btm(BLACK, "thin")
    row += 1

    ws.row_dimensions[row].height = 10
    row += 1

    # ════════════════════════════════════════
    # АДРЕСА: подрядчик | клиент
    # ════════════════════════════════════════
    contractor = [
        ("Qirim LLC", True),
        ("1105 E Orange Grove Blvd, Pasadena, CA 91104", False),
        ("adam@handyeddy.com", False),
        ("323 763 8686", False),
    ]
    client_info = []
    dirty = {"n/a", "na", ""}
    if client_val   and client_val.lower()   not in dirty: client_info.append((client_val,   True))
    if location_val and location_val.lower() not in dirty: client_info.append((location_val, False))
    if phone_val    and phone_val.lower()    not in dirty: client_info.append((phone_val,    False))
    if email_val    and email_val.lower()    not in dirty: client_info.append((email_val,    False))

    for i in range(max(len(contractor), len(client_info))):
        ws.row_dimensions[row].height = 16
        if i < len(contractor):
            txt, bold = contractor[i]
            ws[f"A{row}"]           = txt
            ws[f"A{row}"].font      = fnt(10, bold=bold)
            ws[f"A{row}"].alignment = L
        if i < len(client_info):
            txt, bold = client_info[i]
            ws[f"B{row}"]           = txt
            ws[f"B{row}"].font      = fnt(10, bold=bold)
            ws[f"B{row}"].alignment = L
        row += 1

    ws.row_dimensions[row].height = 15
    row += 1

    # ════════════════════════════════════════
    # ТАБЛИЦА LABOR
    # ════════════════════════════════════════

    # Заголовок (С border — оставляем)
    ws.row_dimensions[row].height = 24
    ws[f"A{row}"]           = "Description"
    ws[f"A{row}"].font      = fnt(10, bold=True)
    ws[f"A{row}"].fill      = header_fill
    ws[f"A{row}"].alignment = L
    ws[f"A{row}"].border    = full()
    ws[f"B{row}"]           = "Labor cost"
    ws[f"B{row}"].font      = fnt(10, bold=True)
    ws[f"B{row}"].fill      = header_fill
    ws[f"B{row}"].alignment = R
    ws[f"B{row}"].border    = full()
    row += 1

    # Labor строки — БЕЗ borders (только зебра-заливка)
    labor_start = row
    for i, (desc, cost_val) in enumerate(labor_items):
        lines = max(1, len(desc) // 70 + 1)
        ws.row_dimensions[row].height = max(20, lines * 14)
        fill = gray_fill if i % 2 == 0 else white_fill

        ws[f"A{row}"]               = f"- {desc}"
        ws[f"A{row}"].font          = fnt(10)
        ws[f"A{row}"].fill          = fill
        ws[f"A{row}"].alignment     = LW
        ws[f"B{row}"]               = cost_val
        ws[f"B{row}"].font          = fnt(10)
        ws[f"B{row}"].fill          = fill
        ws[f"B{row}"].alignment     = RT
        ws[f"B{row}"].number_format = "$#,##0.00"
        row += 1
    labor_end = row - 1

    ws.row_dimensions[row].height = 8
    row += 1

    # ════════════════════════════════════════
    # ТАБЛИЦА MATERIALS (если есть)
    # ════════════════════════════════════════
    mat_start = None
    mat_end   = None

    if material_items:
        # Заголовок (С border)
        ws.row_dimensions[row].height = 24
        ws[f"A{row}"]           = "Materials"
        ws[f"A{row}"].font      = fnt(10, bold=True)
        ws[f"A{row}"].fill      = header_fill
        ws[f"A{row}"].alignment = L
        ws[f"A{row}"].border    = full()
        ws[f"B{row}"]           = "Cost"
        ws[f"B{row}"].font      = fnt(10, bold=True)
        ws[f"B{row}"].fill      = header_fill
        ws[f"B{row}"].alignment = R
        ws[f"B{row}"].border    = full()
        row += 1

        # Materials строки — БЕЗ borders
        mat_start = row
        for i, (desc, cost_val) in enumerate(material_items):
            lines = max(1, len(desc) // 70 + 1)
            ws.row_dimensions[row].height = max(20, lines * 14)
            fill = gray_fill if i % 2 == 0 else white_fill

            ws[f"A{row}"]               = f"- {desc}"
            ws[f"A{row}"].font          = fnt(10)
            ws[f"A{row}"].fill          = fill
            ws[f"A{row}"].alignment     = LW
            ws[f"B{row}"]               = cost_val
            ws[f"B{row}"].font          = fnt(10)
            ws[f"B{row}"].fill          = fill
            ws[f"B{row}"].alignment     = RT
            ws[f"B{row}"].number_format = "$#,##0.00"
            row += 1
        mat_end = row - 1

        ws.row_dimensions[row].height = 8
        row += 1

    # ════════════════════════════════════════
    # ИТОГИ С ФОРМУЛАМИ
    # ════════════════════════════════════════

    # Subtotal Labor
    subtotal_labor_row = row
    ws.row_dimensions[row].height = 22
    ws[f"A{row}"]               = "Subtotal Labor"
    ws[f"A{row}"].font          = fnt(11, bold=True)
    ws[f"A{row}"].alignment     = L
    ws[f"A{row}"].border        = med()
    ws[f"B{row}"]               = f"=SUM(B{labor_start}:B{labor_end})"
    ws[f"B{row}"].font          = fnt(11, bold=True)
    ws[f"B{row}"].alignment     = R
    ws[f"B{row}"].number_format = "$#,##0.00"
    ws[f"B{row}"].border        = med()
    row += 1

    # Estimated Cost of Materials
    subtotal_mat_row = row
    ws.row_dimensions[row].height = 22
    ws[f"A{row}"]               = "Estimated Cost of Materials"
    ws[f"A{row}"].font          = fnt(11, bold=True)
    ws[f"A{row}"].alignment     = L
    ws[f"A{row}"].border        = med()
    if mat_start and mat_end:
        ws[f"B{row}"] = f"=SUM(B{mat_start}:B{mat_end})"
    else:
        ws[f"B{row}"] = 0
    ws[f"B{row}"].font          = fnt(11, bold=True)
    ws[f"B{row}"].alignment     = R
    ws[f"B{row}"].number_format = "$#,##0.00"
    ws[f"B{row}"].border        = med()
    row += 1

    # Discount — РЕДАКТИРУЕМЫЙ ПРОЦЕНТ через скрытую ячейку C
    # Адам меняет число в C — всё пересчитывается автоматически
    # Discount применяется ТОЛЬКО к Subtotal Labor (не к materials)
    discount_row = row
    ws.row_dimensions[row].height = 22
    ws[f"C{row}"]               = 0.10  # ← редактируемый процент (10% по умолчанию)
    ws[f"A{row}"]               = f'="Discount %"&TEXT(C{row}*100,"0")'
    ws[f"A{row}"].font          = fnt(11, bold=True)
    ws[f"A{row}"].alignment     = L
    ws[f"A{row}"].border        = med()
    ws[f"B{row}"]               = f"=-C{row}*B{subtotal_labor_row}"
    ws[f"B{row}"].font          = fnt(11, bold=True)
    ws[f"B{row}"].alignment     = R
    ws[f"B{row}"].number_format = "$#,##0.00"
    ws[f"B{row}"].border        = med()
    row += 1

    # TOTAL
    ws.row_dimensions[row].height = 26
    ws[f"A{row}"]               = "TOTAL"
    ws[f"A{row}"].font          = fnt(12, bold=True, underline=True)
    ws[f"A{row}"].alignment     = L
    ws[f"A{row}"].border        = dbl()
    ws[f"B{row}"]               = f"=B{subtotal_labor_row}+B{subtotal_mat_row}+B{discount_row}"
    ws[f"B{row}"].font          = fnt(12, bold=True, underline=True)
    ws[f"B{row}"].alignment     = R
    ws[f"B{row}"].number_format = "$#,##0.00"
    ws[f"B{row}"].border        = dbl()
    row += 2

    # ════════════════════════════════════════
    # TIMELINE
    # ════════════════════════════════════════
    if timeline_val:
        ws.row_dimensions[row].height = 18
        ws.merge_cells(f"A{row}:B{row}")
        ws[f"A{row}"]           = f"Estimated Timeline: {timeline_val}"
        ws[f"A{row}"].font      = fnt(10, color=DARK_GRAY)
        ws[f"A{row}"].alignment = L
        row += 2

    # ════════════════════════════════════════
    # NOTES (только из текста агента)
    # ════════════════════════════════════════
    if notes_lines:
        ws.row_dimensions[row].height = 18
        ws.merge_cells(f"A{row}:B{row}")
        ws[f"A{row}"]           = "Notes:"
        ws[f"A{row}"].font      = fnt(10, bold=True)
        ws[f"A{row}"].alignment = L
        row += 1

        for note in notes_lines:
            lines = max(1, len(note) // 90 + 1)
            ws.row_dimensions[row].height = max(18, lines * 14)
            ws.merge_cells(f"A{row}:B{row}")
            ws[f"A{row}"]           = f"- {note}"
            ws[f"A{row}"].font      = fnt(9, color=DARK_GRAY)
            ws[f"A{row}"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            row += 1

    row += 1

    # ════════════════════════════════════════
    # FOOTER
    # ════════════════════════════════════════
    ws.row_dimensions[row].height = 3
    for col in ["A", "B"]:
        ws[f"{col}{row}"].border = btm(BLACK, "thin")
    row += 1

    ws.row_dimensions[row].height = 16
    ws.merge_cells(f"A{row}:B{row}")
    ws[f"A{row}"]           = "Qirim LLC  |  adam@handyeddy.com  |  323 763 8686"
    ws[f"A{row}"].font      = fnt(9, color=DARK_GRAY)
    ws[f"A{row}"].alignment = Alignment(horizontal="center", vertical="center")

    ws.print_area = f"A1:B{row}"
    wb.save(filepath)
    return filepath
