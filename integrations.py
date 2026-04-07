import os
import requests
import dropbox
import sendgrid
import base64
import json
from dropbox.files import WriteMode
from dropbox.exceptions import ApiError
from dotenv import load_dotenv
from datetime import datetime
from sendgrid.helpers.mail import Mail, Attachment, FileContent, FileName, FileType, Disposition

load_dotenv()

DROPBOX_TOKEN = os.getenv("DROPBOX_TOKEN")
SENDGRID_API_KEY = os.getenv("SENDGRID_API_KEY")
SENDGRID_FROM_EMAIL = os.getenv("SENDGRID_FROM_EMAIL")
MONDAY_API_KEY = os.getenv("MONDAY_API_KEY")

MONDAY_ESTIMATES_BOARD = 18406712174
MONDAY_CONTACTS_BOARD = 7023159061
DROPBOX_ROOT = "/Aydin Projects 2026"

def get_folder_name(dbx, client_name: str) -> str:
    date_str = datetime.now().strftime("%m%d%y")
    try:
        result = dbx.files_list_folder(DROPBOX_ROOT)
        today_folders = [e.name for e in result.entries if e.name.startswith(date_str)]
        seq = 11 + len(today_folders)
    except:
        seq = 11
    return f"{date_str}-{seq}-{client_name}"

def upload_to_dropbox(pdf_path: str, client_name: str, job_type: str) -> str:
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

def send_estimate_email(client_email: str, client_name: str, job_type: str, estimate_text: str, pdf_path: str, dropbox_url: str) -> bool:
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

def create_monday_cards(client_name: str, job_type: str, estimate_total: str, dropbox_url: str, client_email: str = "", client_phone: str = "", location: str = "") -> bool:
    try:
        headers = {
            "Authorization": MONDAY_API_KEY,
            "Content-Type": "application/json"
        }

        # Estimates board
        col_estimates = json.dumps({"text6__1": job_type, "text__1": dropbox_url})
        query1 = f"""mutation {{
            create_item (
                board_id: {MONDAY_ESTIMATES_BOARD},
                item_name: "{client_name}",
                column_values: {json.dumps(col_estimates)}
            ) {{ id }}
        }}"""
        r1 = requests.post("https://api.monday.com/v2", headers=headers, json={"query": query1})
        data1 = r1.json()
        if "errors" in data1:
            print(f"❌ Monday Estimates: {data1['errors']}")
        else:
            print(f"✅ Estimates карточка создана")

        # Contacts board
        col_contacts = json.dumps({
            "email__1": {"email": client_email, "text": client_email} if client_email else {},
            "phone__1": {"phone": client_phone, "countryShortName": "US"} if client_phone else {},
            "text__1": location,
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
