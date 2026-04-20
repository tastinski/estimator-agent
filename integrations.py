import os
import requests
import dropbox
import sendgrid
import base64
import json
import traceback
import io
import re as _re
from PIL import Image
from dropbox.files import WriteMode
from dotenv import load_dotenv
from datetime import datetime
from sendgrid.helpers.mail import (
    Mail, Attachment, FileContent,
    FileName, FileType, Disposition
)

load_dotenv()

DROPBOX_APP_KEY       = os.getenv("DROPBOX_APP_KEY")
DROPBOX_APP_SECRET    = os.getenv("DROPBOX_APP_SECRET")
DROPBOX_REFRESH_TOKEN = os.getenv("DROPBOX_REFRESH_TOKEN")
SENDGRID_API_KEY      = os.getenv("SENDGRID_API_KEY")
SENDGRID_FROM_EMAIL   = os.getenv("SENDGRID_FROM_EMAIL")
MONDAY_API_KEY        = os.getenv("MONDAY_API_KEY")

MONDAY_MAIN_BOARD     = 18393766403   # Main board (новый)
MONDAY_CONTACTS_BOARD = 7023159061    # Contacts board (не менялся)
MONDAY_FAST_GROUP     = "new_group__1"  # Fast projects Estimate
DROPBOX_ROOT          = "/Aydin Projects 2026"

DIRTY_VALUES = ["na", "n/a", "unknown", "needed", "[needed]",
                "none", "pending", "[pending]", "n\\a", ""]


# ─────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────

def _normalize(s):
    """Убираем точки, тире, подчёркивания, лишние пробелы для нечёткого поиска"""
    if not s:
        return ""
    return _re.sub(r"[.\-_]", " ", s).lower().strip()


def _clean_phone(raw):
    if not raw:
        return ""
    cleaned = (raw.replace("(", "").replace(")", "")
                  .replace(" ", "").replace("-", "").strip())
    return "" if cleaned.lower() in DIRTY_VALUES else cleaned


def _clean_email(raw):
    if not raw:
        return ""
    return "" if raw.strip().lower() in DIRTY_VALUES else raw.strip()


def _phone_digits_only(raw):
    """Только цифры — для сравнения телефонов между собой"""
    if not raw:
        return ""
    return _re.sub(r"\D", "", raw)


def get_dbx():
    return dropbox.Dropbox(
        oauth2_refresh_token=DROPBOX_REFRESH_TOKEN,
        app_key=DROPBOX_APP_KEY,
        app_secret=DROPBOX_APP_SECRET
    )


def list_all_dropbox_entries(dbx, path):
    try:
        result = dbx.files_list_folder(path)
        entries = list(result.entries)
        while result.has_more:
            result = dbx.files_list_folder_continue(result.cursor)
            entries.extend(result.entries)
        return entries
    except Exception as e:
        print(f"list_all_dropbox_entries error ({path}): {e}")
        return []


def compress_image_for_claude(photo_data, max_size=1568):
    try:
        img = Image.open(io.BytesIO(photo_data))
        if img.mode != "RGB":
            img = img.convert("RGB")
        img.thumbnail((max_size, max_size))
        buffer = io.BytesIO()
        img.save(buffer, format="JPEG", quality=80)
        return buffer.getvalue()
    except Exception as e:
        print(f"compress_image_for_claude error: {e}")
        return photo_data


# ─────────────────────────────────────────
# NUMBER GENERATION
# ─────────────────────────────────────────

def _extract_number_from_folder(folder_name):
    """Из имени папки '26042011 Maria Gonzalez' вытащить '26042011'"""
    match = _re.match(r"^(\d{8})", folder_name)
    return match.group(1) if match else None


def _extract_name_from_folder(folder_name):
    """Из '26042011 Maria Gonzalez' вытащить 'Maria Gonzalez'"""
    match = _re.match(r"^\d{8}\s*(.+)$", folder_name)
    if match:
        return match.group(1).strip()
    # fallback для старых папок типа "041326-13-Maria Gonzalez"
    parts = _re.split(r"[-\s]+", folder_name, maxsplit=3)
    return parts[-1].strip() if parts else folder_name


def _generate_new_number(all_entries):
    """
    Генерим YYMMDDNN — сегодняшняя дата + порядковый номер.
    NN начинается с 11, инкрементится для каждой новой папки в этот день.
    """
    today = datetime.now().strftime("%y%m%d")  # '260420'
    count_today = 0
    for e in all_entries:
        if not hasattr(e, "name"):
            continue
        num = _extract_number_from_folder(e.name)
        if num and num.startswith(today):
            count_today += 1
    seq = 11 + count_today
    return f"{today}{seq:02d}"


# ─────────────────────────────────────────
# DROPBOX FOLDER SEARCH
# ─────────────────────────────────────────

def find_client_folder(client_name, file_number=None):
    """
    Ищем папку клиента в Dropbox.
    Приоритет:
      1. По номеру (если передан file_number)
      2. По имени клиента
    Возвращает: (folder_path, folder_name, file_number) или (None, None, None)
    """
    try:
        dbx = get_dbx()

        try:
            dbx.files_get_metadata(DROPBOX_ROOT)
        except dropbox.exceptions.ApiError:
            dbx.files_create_folder_v2(DROPBOX_ROOT)

        all_entries = list_all_dropbox_entries(dbx, DROPBOX_ROOT)

        # 1. По номеру
        if file_number:
            for entry in all_entries:
                if not hasattr(entry, "name"):
                    continue
                num = _extract_number_from_folder(entry.name)
                if num == file_number:
                    print(f"Папка найдена по номеру {file_number}: {entry.name}")
                    return entry.path_display, entry.name, num

        # 2. По имени
        if client_name:
            needle = _normalize(client_name)
            for entry in all_entries:
                if not hasattr(entry, "name"):
                    continue
                folder_name_part = _extract_name_from_folder(entry.name)
                if _normalize(folder_name_part) == needle:
                    num = _extract_number_from_folder(entry.name)
                    print(f"Папка найдена по имени '{client_name}': {entry.name}")
                    return entry.path_display, entry.name, num

        print(f"Папка не найдена для client='{client_name}' number='{file_number}'")
        return None, None, None

    except Exception as e:
        print(f"find_client_folder error: {e}")
        traceback.print_exc()
        return None, None, None


# ─────────────────────────────────────────
# DROPBOX UPLOAD
# ─────────────────────────────────────────

def upload_to_dropbox(file_path, client_name, job_type, file_number=None):
    """
    Загружаем estimate в Dropbox.
    - Ищем существующую папку (по номеру, затем по имени)
    - Нашли — кладём туда (с датой в имени файла чтобы не перезаписать)
    - Не нашли — создаём новую с указанным или новым номером

    Возвращает: (shared_link_url, file_number)
    """
    try:
        dbx = get_dbx()

        try:
            dbx.files_get_metadata(DROPBOX_ROOT)
        except dropbox.exceptions.ApiError:
            dbx.files_create_folder_v2(DROPBOX_ROOT)

        all_entries = list_all_dropbox_entries(dbx, DROPBOX_ROOT)

        # Ищем существующую папку
        folder_path = None
        folder_name = None
        found_number = None

        # 1. По номеру
        if file_number:
            for entry in all_entries:
                if not hasattr(entry, "name"):
                    continue
                num = _extract_number_from_folder(entry.name)
                if num == file_number:
                    folder_path = entry.path_display
                    folder_name = entry.name
                    found_number = num
                    print(f"Найдена папка по номеру: {folder_name}")
                    break

        # 2. По имени
        if not folder_path and client_name:
            needle = _normalize(client_name)
            for entry in all_entries:
                if not hasattr(entry, "name"):
                    continue
                folder_name_part = _extract_name_from_folder(entry.name)
                if _normalize(folder_name_part) == needle:
                    folder_path = entry.path_display
                    folder_name = entry.name
                    found_number = _extract_number_from_folder(entry.name)
                    print(f"Найдена папка по имени: {folder_name}")
                    break

        # 3. Не нашли — создаём новую
        if not folder_path:
            if not file_number:
                file_number = _generate_new_number(all_entries)
            folder_name = f"{file_number} {client_name}"
            folder_path = f"{DROPBOX_ROOT}/{folder_name}"
            try:
                dbx.files_create_folder_v2(folder_path)
                print(f"Создана новая папка: {folder_name}")
            except dropbox.exceptions.ApiError as e:
                if "path/conflict" not in str(e):
                    raise
            found_number = file_number

        # Имя файла с датой+временем — чтобы не перезаписывать предыдущие эстимейты
        timestamp = datetime.now().strftime("%m%d-%H%M")
        file_name = f"{client_name} estimate {timestamp}.xlsx"
        dropbox_path = f"{folder_path}/{file_name}"

        with open(file_path, "rb") as f:
            dbx.files_upload(f.read(), dropbox_path, mode=WriteMode.overwrite)

        url = _get_shared_link(dbx, dropbox_path)
        print(f"Dropbox OK: {dropbox_path}")
        return url, found_number

    except Exception as e:
        print(f"upload_to_dropbox error: {e}")
        traceback.print_exc()
        return "", file_number or ""


def _get_shared_link(dbx, dropbox_path):
    try:
        link = dbx.sharing_create_shared_link_with_settings(dropbox_path)
        return link.url
    except dropbox.exceptions.ApiError as e:
        if "shared_link_already_exists" in str(e):
            links = dbx.sharing_list_shared_links(path=dropbox_path, direct_only=True)
            return links.links[0].url if links.links else dropbox_path
        return dropbox_path


# ─────────────────────────────────────────
# DROPBOX PHOTOS
# ─────────────────────────────────────────

def get_dropbox_photos(client_name_part, file_number=None):
    """Загружаем фото из папки клиента. Приоритет — по номеру, потом по имени."""
    try:
        folder_path, _, _ = find_client_folder(client_name_part, file_number)
        if not folder_path:
            return []

        dbx = get_dbx()
        folder_entries = list_all_dropbox_entries(dbx, folder_path)
        photos = []
        for entry in folder_entries:
            if not hasattr(entry, "name"):
                continue
            if not entry.name.lower().endswith((".jpg", ".jpeg", ".png", ".webp")):
                continue
            _, response = dbx.files_download(entry.path_display)
            compressed = compress_image_for_claude(response.content)
            photo_b64 = base64.standard_b64encode(compressed).decode("utf-8")
            photos.append({"name": entry.name, "data": photo_b64})
            print(f"Загружено фото: {entry.name}")
            if len(photos) >= 10:
                break

        print(f"Всего фото загружено: {len(photos)}")
        return photos

    except Exception as e:
        print(f"get_dropbox_photos error: {e}")
        traceback.print_exc()
        return []


# ─────────────────────────────────────────
# EMAIL
# ─────────────────────────────────────────

def send_estimate_email(client_email, client_name, job_type,
                        estimate_text="", pdf_path=None, dropbox_url=""):
    try:
        sg = sendgrid.SendGridAPIClient(api_key=SENDGRID_API_KEY)
        html = (
            f"<p>Hi {client_name},</p>"
            f"<p>Thank you for choosing Handy Eddy! "
            f"Please find your estimate for <b>{job_type}</b> below.</p>"
        )
        if dropbox_url:
            html += f"<p><a href='{dropbox_url}'>📁 View / Download Estimate (Excel)</a></p>"
        html += "<p>Feel free to reply with any questions.<br>— Handy Eddy Team</p>"

        message = Mail(
            from_email=SENDGRID_FROM_EMAIL,
            to_emails=client_email,
            subject=f"Your Estimate — {job_type}",
            html_content=html
        )

        if pdf_path and os.path.exists(pdf_path):
            with open(pdf_path, "rb") as f:
                pdf_data = base64.b64encode(f.read()).decode()
            attachment = Attachment(
                FileContent(pdf_data),
                FileName(f"estimate_{client_name.replace(' ', '_')}.xlsx"),
                FileType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
                Disposition("attachment")
            )
            message.attachment = attachment

        sg.send(message)
        print(f"Email sent to {client_email} for {client_name}")
        return True

    except Exception as e:
        print(f"send_estimate_email error: {e}")
        return False


# ─────────────────────────────────────────
# MONDAY — FIND EXISTING CLIENT BY PHONE
# ─────────────────────────────────────────

def find_client_by_phone(phone):
    """Ищем существующую карточку по телефону. Возвращает dict или None."""
    phone_digits = _phone_digits_only(phone)
    if not phone_digits or len(phone_digits) < 7:
        return None

    try:
        headers = {"Authorization": MONDAY_API_KEY, "Content-Type": "application/json",
                   "API-Version": "2023-10"}
        query = """query ($board: [ID!]) {
            boards(ids: $board) {
                items_page(limit: 500) {
                    items {
                        id
                        name
                        column_values { id text }
                    }
                }
            }
        }"""
        r = requests.post(
            "https://api.monday.com/v2", headers=headers,
            json={"query": query, "variables": {"board": [str(MONDAY_MAIN_BOARD)]}}
        )
        data = r.json()
        items = (
            data.get("data", {})
                .get("boards", [{}])[0]
                .get("items_page", {})
                .get("items", [])
        )

        for item in items:
            cols = {c["id"]: c["text"] for c in item["column_values"]}
            item_phone_digits = _phone_digits_only(cols.get("phone__1", ""))
            if not item_phone_digits:
                continue
            # Сравниваем последние 10 цифр (убирает проблему с +1 / без +1)
            if item_phone_digits[-10:] == phone_digits[-10:]:
                return {
                    "id":          item["id"],
                    "name":        item["name"],
                    "file_number": cols.get("text5__1", ""),
                    "email":       cols.get("email__1", ""),
                    "phone":       cols.get("phone__1", ""),
                    "job_type":    cols.get("text6__1", ""),
                    "dropbox_url": cols.get("link__1", ""),
                    "location":    cols.get("location__1", ""),
                }

        return None

    except Exception as e:
        print(f"find_client_by_phone error: {e}")
        return None


# ─────────────────────────────────────────
# MONDAY — CREATE CARDS
# ─────────────────────────────────────────

def create_monday_cards(client_name, job_type, estimate_total,
                        dropbox_url, client_email="", client_phone="",
                        location="", file_number=""):
    """
    Создаём карточки в Main и Contacts.
    file_number — если передан, используется в колонке # и имени Dropbox-папки.
    Если не передан — генерим на основе даты.
    """
    try:
        headers = {"Authorization": MONDAY_API_KEY, "Content-Type": "application/json",
                   "API-Version": "2023-10"}
        phone_clean = _clean_phone(client_phone)
        email_val = _clean_email(client_email)

        # Если file_number не передан — генерим новый через Dropbox (консистентно)
        if not file_number:
            try:
                dbx = get_dbx()
                entries = list_all_dropbox_entries(dbx, DROPBOX_ROOT)
                file_number = _generate_new_number(entries)
            except Exception:
                file_number = datetime.now().strftime("%y%m%d") + "11"

        col_main = {
            "text5__1":  file_number,
            "phone__1":  {"phone": phone_clean, "countryShortName": "US"} if phone_clean else {},
            "email__1":  {"email": email_val, "text": email_val} if email_val else {},
            "text6__1":  job_type,
            "link__1":   {"url": dropbox_url, "text": "Estimate XLSX"} if dropbox_url else {},
            "location__1": {"address": location} if location else {},
        }

        query1 = """mutation ($board: ID!, $group: String!, $name: String!, $cols: JSON!) {
            create_item(board_id: $board, group_id: $group,
                        item_name: $name, column_values: $cols) { id }
        }"""
        r1 = requests.post(
            "https://api.monday.com/v2", headers=headers,
            json={"query": query1, "variables": {
                "board": MONDAY_MAIN_BOARD,
                "group": MONDAY_FAST_GROUP,
                "name":  client_name,
                "cols":  json.dumps(col_main)
            }}
        )
        data1 = r1.json()
        if "errors" in data1:
            return f"MAIN_ERROR: {data1['errors']}"
        main_id = data1.get("data", {}).get("create_item", {}).get("id", "NO_ID")

        col_contacts = {
            "phone":       {"phone": phone_clean, "countryShortName": "US"} if phone_clean else {},
            "email":       {"email": email_val, "text": email_val} if email_val else {},
            "location__1": {"address": location} if location else {},
            "text4":       job_type,
        }
        query2 = """mutation ($board: ID!, $name: String!, $cols: JSON!) {
            create_item(board_id: $board, item_name: $name, column_values: $cols) { id }
        }"""
        r2 = requests.post(
            "https://api.monday.com/v2", headers=headers,
            json={"query": query2, "variables": {
                "board": MONDAY_CONTACTS_BOARD,
                "name":  client_name,
                "cols":  json.dumps(col_contacts)
            }}
        )
        data2 = r2.json()
        if "errors" in data2:
            return f"CONTACTS_ERROR: {data2['errors']}"
        contacts_id = data2.get("data", {}).get("create_item", {}).get("id", "NO_ID")

        return f"OK: main={main_id} contacts={contacts_id} number={file_number}"

    except Exception as e:
        print(f"create_monday_cards error: {e}")
        traceback.print_exc()
        return False


# ─────────────────────────────────────────
# MONDAY — FIND CLIENT BY NAME
# ─────────────────────────────────────────

def find_client_in_monday(client_name_query):
    try:
        headers = {"Authorization": MONDAY_API_KEY, "Content-Type": "application/json",
                   "API-Version": "2023-10"}
        query = """query ($board: [ID!]) {
            boards(ids: $board) {
                items_page(limit: 500) {
                    items {
                        id
                        name
                        column_values { id text }
                    }
                }
            }
        }"""
        r = requests.post(
            "https://api.monday.com/v2", headers=headers,
            json={"query": query, "variables": {"board": [str(MONDAY_MAIN_BOARD)]}}
        )
        data = r.json()
        items = (
            data.get("data", {})
                .get("boards", [{}])[0]
                .get("items_page", {})
                .get("items", [])
        )

        query_lower = client_name_query.lower().strip()
        matches = []
        for item in items:
            if query_lower in item["name"].lower():
                cols = {c["id"]: c["text"] for c in item["column_values"]}
                matches.append({
                    "id":          item["id"],
                    "name":        item["name"],
                    "file_number": cols.get("text5__1", ""),
                    "email":       cols.get("email__1", ""),
                    "phone":       cols.get("phone__1", ""),
                    "job_type":    cols.get("text6__1", ""),
                    "dropbox_url": cols.get("link__1", ""),
                    "location":    cols.get("location__1", ""),
                })

        if len(matches) == 0:
            return None
        if len(matches) == 1:
            return matches[0]
        return matches

    except Exception as e:
        print(f"find_client_in_monday error: {e}")
        return None


# ─────────────────────────────────────────
# MONDAY — UPDATE LINK
# ─────────────────────────────────────────

def update_monday_item_link(item_id, dropbox_url):
    try:
        headers = {"Authorization": MONDAY_API_KEY, "Content-Type": "application/json",
                   "API-Version": "2023-10"}
        query = """mutation ($item: ID!, $board: ID!, $col: String!, $val: JSON!) {
            change_column_value(item_id: $item, board_id: $board,
                                column_id: $col, value: $val) { id }
        }"""
        val = json.dumps({"url": dropbox_url, "text": "Estimate XLSX"})
        r = requests.post(
            "https://api.monday.com/v2", headers=headers,
            json={"query": query, "variables": {
                "item":  str(item_id),
                "board": str(MONDAY_MAIN_BOARD),
                "col":   "link__1",
                "val":   val
            }}
        )
        print(f"update_monday_item_link {item_id}: {r.status_code} {r.text[:200]}")
    except Exception as e:
        print(f"update_monday_item_link error: {e}")
