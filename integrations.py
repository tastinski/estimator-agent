import os
import requests
import dropbox
import sendgrid
import base64
import json
from dropbox.files import WriteMode
from dotenv import load_dotenv
from datetime import datetime
from sendgrid.helpers.mail import Mail, Attachment, FileContent, FileName, FileType, Disposition

load_dotenv()

DROPBOX_TOKEN = os.getenv("DROPBOX_TOKEN")
SENDGRID_API_KEY = os.getenv("SENDGRID_API_KEY")
SENDGRID_FROM_EMAIL = os.getenv("SENDGRID_FROM_EMAIL")
MONDAY_API_KEY = os.getenv("MONDAY_API_KEY")

MONDAY_MAIN_BOARD = 18406712174
MONDAY_CONTACTS_BOARD = 7023159061
DROPBOX_ROOT = "/Aydin Projects 2026"

def get_folder_name(dbx, client_name):
    date_str = datetime.now().strftime("%m%d%y")
    try:
        result = dbx.files_list_folder(DROPBOX_ROOT)
        today_folders = [e.name for e in result.entries if e.name.startswith(date_str)]
        seq = 11 + len(today_folders)
    except:
        seq = 11
    return f"{date_str}-{seq}-{client_name}"

def upload_to_dropbox(pdf_path, client_name, job_type):
    try:
        dbx = dropbox.Dropbox(DROPBOX_TOKEN)
        folder_name = get_folder_name(dbx, client_name)
        dropbox_folder = f"{DROPBOX_ROOT}/{folder_name}"
        pdf_name = f"{folder_name} estimate.pdf"
        dropbox_path = f"{dropbox_folder}/{pdf_name}"
        with open(pdf_path, 'rb') as f:
            dbx.files_upload(f.read(), dropbox_path, mode=WriteMode.overwrite)
        try:
            link = dbx.sharing_create_shared_link(dropbox_path)
            shared_url = link.url
        except Exception:
            links = dbx.sharing_list_shared_links(path=dropbox_path)
            shared_url = links.links[0].url if links.links else dropbox_path
        print(f"✅ Dropbox: {dropbox_path}")
        return shared_url
    except Exception as e:
        print(f"❌ Dropbox ошибка: {e}")
        return ""

def send_estimate_email(client_email, client_name, job_type, estimate_text, pdf_path, dropbox_url):
    try:
        sg = sendgrid.SendGridAPIClient(api_key=SENDGRID_API_KEY)
        with open(pdf_path, 'rb') as f:
            pdf_data = base64.b64encode(f.read()).decode()
        message = Mail(
            from_email=SENDGRID_FROM_EMAIL,
            to_emails=client_email,
            subject=f"Your Estimate — {job_type}",
            html_content=f"<p>Hi {client_name},</p><p>Please find your estimate attached.</p><p><a href='{dropbox_url}'>View in Dropbox</a></p>"
        )
        from sendgrid.helpers.mail import Attachment, FileContent, FileName, FileType, Disposition
        attachment = Attachment(
            FileContent(pdf_data),
            FileName(f"estimate_{client_name.replace(' ', '_')}.pdf"),
            FileType("application/pdf"),
            Disposition("attachment")
        )
        message.attachment = attachment
        sg.send(message)
        return True
    except Exception as e:
        print(f"❌ Email ошибка: {e}")
        return False

def create_monday_cards(client_name, job_type, estimate_total, dropbox_url, client_email="", client_phone="", location=""):
    try:
        headers = {
            "Authorization": MONDAY_API_KEY,
            "Content-Type": "application/json"
        }

        # Карточка 1 — Aydin Main 2026 (Fast projects Estimate)
        col_main = json.dumps({
            "phone__1": {"phone": client_phone, "countryShortName": "US"} if client_phone else {},
            "email__1": {"email": client_email, "text": client_email} if client_email else {},
            "text6__1": job_type,
            "link__1": {"url": dropbox_url, "text": "Estimate PDF"} if dropbox_url else {},
            "location__1": location,
        })
        query1 = f"""mutation {{
            create_item (
                board_id: {MONDAY_MAIN_BOARD},
                group_id: "group_mm0ea2gw",
                item_name: "{client_name}",
                column_values: {json.dumps(col_main)}
            ) {{ id }}
        }}"""
        r1 = requests.post("https://api.monday.com/v2", headers=headers, json={"query": query1})
        data1 = r1.json()
        if "errors" in data1:
            print(f"❌ Monday Main: {data1['errors']}")
        else:
            print(f"✅ Fast projects Estimate карточка создана")

        # Карточка 2 — Contacts
        col_contacts = json.dumps({
            "phone": {"phone": client_phone, "countryShortName": "US"} if client_phone else {},
            "email": {"email": client_email, "text": client_email} if client_email else {},
            "location__1": location,
            "text4": job_type,
        })
        query2 = f"""mutation {{
            create_item (
                board_id: {MONDAY_CONTACTS_BOARD},
                item_name: "{client_name}",
                column_values: {json.dumps(col_contacts)}
            ) {{ id }}
        }}"""
        r2 = requests.post("https://api.monday.com/v2", headers=headers, json={"query": query2})
        data2 = r2.json()
        if "errors" in data2:
            print(f"❌ Monday Contacts: {data2['errors']}")
        else:
            print(f"✅ Contacts карточка создана")

        return True
    except Exception as e:
        print(f"❌ Ошибка: {e}")
        return False
