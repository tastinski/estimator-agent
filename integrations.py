import os
import requests
import dropbox
import sendgrid
import base64
import json
import traceback
import io
from PIL import Image
from dropbox.files import WriteMode
from dotenv import load_dotenv
from datetime import datetime
from sendgrid.helpers.mail import (
    Mail, Attachment, FileContent,
    FileName, FileType, Disposition
)

load_dotenv()

DROPBOX_APP_KEY      = os.getenv("DROPBOX_APP_KEY")
DROPBOX_APP_SECRET   = os.getenv("DROPBOX_APP_SECRET")
DROPBOX_REFRESH_TOKEN = os.getenv("DROPBOX_REFRESH_TOKEN")
SENDGRID_API_KEY     = os.getenv("SENDGRID_API_KEY")
SENDGRID_FROM_EMAIL  = os.getenv("SENDGRID_FROM_EMAIL")
MONDAY_API_KEY       = os.getenv("MONDAY_API_KEY")

MONDAY_MAIN_BOARD     = 18406712174
MONDAY_CONTACTS_BOARD = 7023159061
DROPBOX_ROOT          = "/Aydin Projects 2026"


# ─────────────────────────────────────────
# DROPBOX HELPERS
# ─────────────────────────────────────────

def get_dbx():
    return dropbox.Dropbox(
        oauth2_refresh_token=DROPBOX_REFRESH_TOKEN,
        app_key=DROPBOX_APP_KEY,
        app_secret=DROPBOX_APP_SECRET
    )


def list_all_dropbox_entries(dbx, path):
    """List ALL entries in a folder with pagination."""
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
        return photo_data  # возвращаем оригинал если PIL не справился


# ─────────────────────────────────────────
# DROPBOX UPLOAD
# ─────────────────────────────────────────

def upload_to_dropbox(file_path, client_name, job_type, existing_folder=None):
    try:
        dbx = get_dbx()

        # Убеждаемся что корневая папка существует
        try:
            dbx.files_get_metadata(DROPBOX_ROOT)
        except dropbox.exceptions.ApiError:
            dbx.files_create_folder_v2(DROPBOX_ROOT)

        all_entries = list_all_dropbox_entries(dbx, DROPBOX_ROOT)

        # Если передана папка клиента — ищем её среди всех папок
        if existing_folder:
            for entry in all_entries:
                if hasattr(entry, "name") and existing_folder.lower() in entry.name.lower():
                    folder_name   = entry.name
                    dropbox_folder = entry.path_display
                    file_name     = f"{folder_name} estimate.xlsx"
                    dropbox_path  = f"{dropbox_folder}/{file_name}"
                    with open(file_path, "rb") as f:
                        dbx.files_upload(f.read(), dropbox_path, mode=WriteMode.overwrite)
                    return _get_shared_link(dbx, dropbox_path)

        # Генерируем новую папку
        date_str = datetime.now().strftime("%m%d%y")

        # Счётчик — берём ВСЕ папки, не только сегодняшние
        existing_count = sum(1 for e in all_entries if hasattr(e, "name"))
        seq = 11 + existing_count

        # Имя карточки с датой-ID: "Maria Gonzalez 260414"
        card_id     = datetime.now().strftime("%y%m%d")
        folder_name  = f"{date_str}-{seq}-{client_name}"
        dropbox_folder = f"{DROPBOX_ROOT}/{folder_name}"
        file_name   = f"{folder_name} estimate.xlsx"
        dropbox_path = f"{dropbox_folder}/{file_name}"

        try:
            dbx.files_create_folder_v2(dropbox_folder)
        except dropbox.exceptions.ApiError as e:
            if "path/conflict" not in str(e):
                raise

        with open(file_path, "rb") as f:
            dbx.files_upload(f.read(), dropbox_path, mode=WriteMode.overwrite)

        url = _get_shared_link(dbx, dropbox_path)
        print(f"Dropbox OK: {dropbox_path}")
        return url

    except Exception as e:
        print(f"upload_to_dropbox error: {e}")
        traceback.print_exc()
        return ""


def _get_shared_link(dbx, dropbox_path):
    """Create or retrieve shared link for a file."""
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

def get_dropbox_photos(client_name_part):
    try:
        dbx = get_dbx()
        all_entries = list_all_dropbox_entries(dbx, DROPBOX_ROOT)

        client_folder = None
        for entry in all_entries:
            if hasattr(entry, "name") and client_name_part.lower() in entry.name.lower():
                client_folder = entry.path_display
                break

        if not client_folder:
            print(f"Папка не найдена для: {client_name_part}")
            return []

        folder_entries = list_all_dropbox_entries(dbx, client_folder)
        photos = []
        for entry in folder_entries:
            if not hasattr(entry, "name"):
                continue
            if not entry.name.lower().endswith((".jpg", ".jpeg", ".png", ".webp")):
                continue
            _, response = dbx.files_download(entry.path_display)
            compressed = compress_image_for_claude(response.content)
            photo_b64  = base64.standard_b64encode(compressed).decode("utf-8")
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

        # Аттач Excel если есть локальный файл
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
# MONDAY — CREATE CARDS
# ─────────────────────────────────────────

def create_monday_cards(client_name, job_type, estimate_total,
                        dropbox_url, client_email="", client_phone="", location=""):
    try:
        headers = {"Authorization": MONDAY_API_KEY, "Content-Type": "application/json"}

        # Имя карточки с уникальным ID: "Maria Gonzalez 260414"
        card_id   = datetime.now().strftime("%y%m%d")
        full_name = f"{client_name} {card_id}"

        phone_clean = (
            client_phone
            .replace("(", "").replace(")", "")
            .replace(" ", "").replace("-", "")
            if client_phone else ""
        )
        if phone_clean.lower() in ["na", "n/a", "unknown"]:
            phone_clean = ""

        email_val = client_email if client_email and client_email.lower() not in ["n/a", "na", ""] else ""

        col_main = {
            "phone__1": {"phone": phone_clean, "countryShortName": "US"} if phone_clean else {},
            "email__1": {"email": email_val, "text": email_val} if email_val else {},
            "text6__1": job_type,
            "link__1":  {"url": dropbox_url, "text": "Estimate XLSX"} if dropbox_url else {},
            "text__1":  location,
        }

        query1 = """mutation ($board: ID!, $group: String!, $name: String!, $cols: JSON!) {
            create_item(board_id: $board, group_id: $group,
                        item_name: $name, column_values: $cols) { id }
        }"""
        r1 = requests.post(
            "https://api.monday.com/v2", headers=headers,
            json={"query": query1, "variables": {
                "board": MONDAY_MAIN_BOARD,
                "group": "new_group__1",
                "name":  full_name,
                "cols":  json.dumps(col_main)
            }}
        )
        data1 = r1.json()
        if "errors" in data1:
            return f"MAIN_ERROR: {data1['errors']}"
        main_id = data1.get("data", {}).get("create_item", {}).get("id", "NO_ID")

        col_contacts = {
            "phone": {"phone": phone_clean, "countryShortName": "US"} if phone_clean else {},
            "email": {"email": email_val, "text": email_val} if email_val else {},
            "text__1": location,
            "text4":   job_type,
        }
        query2 = """mutation ($board: ID!, $name: String!, $cols: JSON!) {
            create_item(board_id: $board, item_name: $name, column_values: $cols) { id }
        }"""
        r2 = requests.post(
            "https://api.monday.com/v2", headers=headers,
            json={"query": query2, "variables": {
                "board": MONDAY_CONTACTS_BOARD,
                "name":  full_name,
                "cols":  json.dumps(col_contacts)
            }}
        )
        data2 = r2.json()
        if "errors" in data2:
            return f"CONTACTS_ERROR: {data2['errors']}"
        contacts_id = data2.get("data", {}).get("create_item", {}).get("id", "NO_ID")

        return f"OK: main={main_id} contacts={contacts_id}"

    except Exception as e:
        print(f"create_monday_cards error: {e}")
        traceback.print_exc()
        return False


# ─────────────────────────────────────────
# MONDAY — FIND CLIENT
# ─────────────────────────────────────────

def find_client_in_monday(client_name_query):
    """
    Search Monday main board by name.
    Returns:
      - None          если не найден
      - dict          если найден ровно один
      - list of dict  если найдено несколько
    """
    try:
        headers = {"Authorization": MONDAY_API_KEY, "Content-Type": "application/json"}

        # Берём до 500 карточек чтобы не промахнуться
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
                    "email":       cols.get("email__1", ""),
                    "phone":       cols.get("phone__1", ""),
                    "job_type":    cols.get("text6__1", ""),
                    "dropbox_url": cols.get("link__1", ""),
                    "location":    cols.get("text__1", ""),
                })

        if len(matches) == 0:
            return None
        if len(matches) == 1:
            return matches[0]
        return matches

    except Exception as e:
        print(f"find_client_in_monday error: {e}")
        return None
