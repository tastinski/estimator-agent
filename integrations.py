import os
import requests
import dropbox
import sendgrid
import base64
import json
import traceback
from dropbox.files import WriteMode
from dotenv import load_dotenv
from datetime import datetime
from sendgrid.helpers.mail import Mail, Attachment, FileContent, FileName, FileType, Disposition

load_dotenv()

DROPBOX_APP_KEY = os.getenv("DROPBOX_APP_KEY")
DROPBOX_APP_SECRET = os.getenv("DROPBOX_APP_SECRET")
DROPBOX_REFRESH_TOKEN = os.getenv("DROPBOX_REFRESH_TOKEN")
SENDGRID_API_KEY = os.getenv("SENDGRID_API_KEY")
SENDGRID_FROM_EMAIL = os.getenv("SENDGRID_FROM_EMAIL")
MONDAY_API_KEY = os.getenv("MONDAY_API_KEY")

MONDAY_MAIN_BOARD = 18406712174
MONDAY_CONTACTS_BOARD = 7023159061
DROPBOX_ROOT = "/Aydin Projects 2026"

def upload_to_dropbox(file_path, client_name, job_type, existing_folder=None):
    try:
        dbx = dropbox.Dropbox(
            oauth2_refresh_token=DROPBOX_REFRESH_TOKEN,
            app_key=DROPBOX_APP_KEY,
            app_secret=DROPBOX_APP_SECRET
        )
if existing_folder:
            try:
                result = dbx.files_list_folder(DROPBOX_ROOT)
                for entry in result.entries:
                    if hasattr(entry, "name") and existing_folder.lower() in entry.name.lower():
                        dropbox_folder = entry.path_display
                        folder_name = entry.name
                        file_name = f"{folder_name} estimate.xlsx"
                        dropbox_path = f"{dropbox_folder}/{file_name}"
                        with open(file_path, "rb") as f:
                            dbx.files_upload(f.read(), dropbox_path, mode=WriteMode.overwrite)
                        try:
                            link = dbx.sharing_create_shared_link_with_settings(dropbox_path)
                            return link.url
                        except dropbox.exceptions.ApiError:
                            links = dbx.sharing_list_shared_links(path=dropbox_path, direct_only=True)
                            return links.links[0].url if links.links else dropbox_path
            except Exception as e:
                print(f"existing_folder error: {e}")
                try:
            dbx.files_get_metadata(DROPBOX_ROOT)
        except dropbox.exceptions.ApiError:
            dbx.files_create_folder_v2(DROPBOX_ROOT)

        date_str = datetime.now().strftime("%m%d%y")
        try:
            result = dbx.files_list_folder(DROPBOX_ROOT)
            today_folders = [e.name for e in result.entries if hasattr(e, "name") and e.name.startswith(date_str)]
            seq = 11 + len(today_folders)
        except Exception as e:
            print(f"list_folder error: {e}")
            seq = 11

        folder_name = f"{date_str}-{seq}-{client_name}"
        dropbox_folder = f"{DROPBOX_ROOT}/{folder_name}"
        file_name = f"{folder_name} estimate.xlsx"
        dropbox_path = f"{dropbox_folder}/{file_name}"

        try:
            dbx.files_create_folder_v2(dropbox_folder)
        except dropbox.exceptions.ApiError as e:
            if "path/conflict" not in str(e):
                raise

        with open(file_path, "rb") as f:
            dbx.files_upload(f.read(), dropbox_path, mode=WriteMode.overwrite)

        try:
            link = dbx.sharing_create_shared_link_with_settings(dropbox_path)
            shared_url = link.url
        except dropbox.exceptions.ApiError as e:
            if "shared_link_already_exists" in str(e):
                links = dbx.sharing_list_shared_links(path=dropbox_path, direct_only=True)
                shared_url = links.links[0].url if links.links else dropbox_path
            else:
                shared_url = dropbox_path

        print(f"Dropbox OK: {dropbox_path}")
        return shared_url

    except Exception as e:
        print(f"Dropbox error: {e}")
        traceback.print_exc()
        return ""

def get_dropbox_photos(client_name_part):
    try:
        dbx = dropbox.Dropbox(
            oauth2_refresh_token=DROPBOX_REFRESH_TOKEN,
            app_key=DROPBOX_APP_KEY,
            app_secret=DROPBOX_APP_SECRET
        )
        result = dbx.files_list_folder(DROPBOX_ROOT)
        client_folder = None
        for entry in result.entries:
            if hasattr(entry, "name") and client_name_part.lower() in entry.name.lower():
                client_folder = entry.path_display
                break
        if not client_folder:
            print(f"Папка не найдена для: {client_name_part}")
            return []
        photos = []
        folder_contents = dbx.files_list_folder(client_folder)
        all_entries = list(folder_contents.entries)
        while folder_contents.has_more:
            folder_contents = dbx.files_list_folder_continue(folder_contents.cursor)
            all_entries.extend(folder_contents.entries)
        for entry in all_entries:
            if hasattr(entry, "name") and entry.name.lower().endswith((".jpg", ".jpeg", ".png", ".webp")):
                _, response = dbx.files_download(entry.path_display)
                photo_data = response.content
                photo_b64 = base64.standard_b64encode(photo_data).decode("utf-8")
                photos.append({"name": entry.name, "data": photo_b64})
                print(f"Загружено фото: {entry.name}")
                if len(photos) >= 10:
                    break
        print(f"Всего фото: {len(photos)}")
        return photos
    except Exception as e:
        print(f"Dropbox фото ошибка: {e}")
        traceback.print_exc()
        return []

def send_estimate_email(client_email, client_name, job_type, estimate_text, pdf_path, dropbox_url):
    try:
        sg = sendgrid.SendGridAPIClient(api_key=SENDGRID_API_KEY)
        with open(pdf_path, "rb") as f:
            pdf_data = base64.b64encode(f.read()).decode()
        message = Mail(
            from_email=SENDGRID_FROM_EMAIL,
            to_emails=client_email,
            subject=f"Your Estimate - {job_type}",
            html_content=f"<p>Hi {client_name},</p><p>Please find your estimate attached.</p><p><a href='{dropbox_url}'>View in Dropbox</a></p>"
        )
        attachment = Attachment(
            FileContent(pdf_data),
            FileName(f"estimate_{client_name.replace(' ', '_')}.xlsx"),
            FileType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
            Disposition("attachment")
        )
        message.attachment = attachment
        sg.send(message)
        return True
    except Exception as e:
        print(f"Email error: {e}")
        return False

def create_monday_cards(client_name, job_type, estimate_total, dropbox_url, client_email="", client_phone="", location=""):
    try:
        headers = {"Authorization": MONDAY_API_KEY, "Content-Type": "application/json"}
        phone_clean = client_phone.replace("(","").replace(")","").replace(" ","").replace("-","") if client_phone else ""
        if phone_clean.lower() in ["na", "n/a", "unknown"]:
            phone_clean = ""
        col_main = {
            "phone__1": {"phone": phone_clean, "countryShortName": "US"} if phone_clean else {},
            "email__1": {"email": client_email, "text": client_email} if client_email and client_email not in ["N/A","n/a",""] else {},
            "text6__1": job_type,
            "link__1": {"url": dropbox_url, "text": "Estimate XLSX"} if dropbox_url else {},
            "text__1": location,
        }
        query1 = """mutation ($board: ID!, $group: String!, $name: String!, $cols: JSON!) {
            create_item(board_id: $board, group_id: $group, item_name: $name, column_values: $cols) { id }
        }"""
        r1 = requests.post("https://api.monday.com/v2", headers=headers, json={"query": query1, "variables": {"board": MONDAY_MAIN_BOARD, "group": "new_group__1", "name": client_name, "cols": json.dumps(col_main)}})
        data1 = r1.json()
        if "errors" in data1:
            return f"MAIN_ERROR: {data1['errors']}"
        main_id = data1.get("data", {}).get("create_item", {}).get("id", "NO_ID")

        col_contacts = {
            "phone": {"phone": phone_clean, "countryShortName": "US"} if phone_clean else {},
            "email": {"email": client_email, "text": client_email} if client_email and client_email not in ["N/A","n/a",""] else {},
            "text__1": location,
            "text4": job_type,
        }
        query2 = """mutation ($board: ID!, $name: String!, $cols: JSON!) {
            create_item(board_id: $board, item_name: $name, column_values: $cols) { id }
        }"""
        r2 = requests.post("https://api.monday.com/v2", headers=headers, json={"query": query2, "variables": {"board": MONDAY_CONTACTS_BOARD, "name": client_name, "cols": json.dumps(col_contacts)}})
        data2 = r2.json()
        if "errors" in data2:
            return f"CONTACTS_ERROR: {data2['errors']}"
        contacts_id = data2.get("data", {}).get("create_item", {}).get("id", "NO_ID")

        return f"OK: main={main_id} contacts={contacts_id}"
    except Exception as e:
        print(f"Monday error: {e}")
        traceback.print_exc()
        return False
