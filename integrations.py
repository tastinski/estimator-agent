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

DROPBOX_APP_KEY = os.getenv("DROPBOX_APP_KEY")
DROPBOX_APP_SECRET = os.getenv("DROPBOX_APP_SECRET")
DROPBOX_REFRESH_TOKEN = os.getenv("DROPBOX_REFRESH_TOKEN")
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
        dbx = dropbox.Dropbox(oauth2_refresh_token=DROPBOX_REFRESH_TOKEN, app_key=DROPBOX_APP_KEY, app_secret=DROPBOX_APP_SECRET)
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

        # Aydin Main 2026 — Fast projects Estimate
        # phone__1=Phone, email__1=Email, text6__1=Job description, link__1=Link, location__1=Location
        col_main = {
            "phone__1": {"phone": client_phone.replace("(", "").replace(")", "").replace(" ", "").replace("-", ""), "countryShortName": "US"} if client_phone else {},
            "email__1": {"email": client_email, "text": client_email} if client_email else {},
            "text6__1": job_type,
            "link__1": {"url": dropbox_url, "text": "Estimate PDF"} if dropbox_url else {},
            "text__1": location,
        }
        query1 = """mutation ($board: ID!, $group: String!, $name: String!, $cols: JSON!) {
            create_item(board_id: $board, group_id: $group, item_name: $name, column_values: $cols) { id }
        }"""
        vars1 = {
            "board": MONDAY_MAIN_BOARD,
            "group": "new_group__1",
            "name": client_name,
            "cols": json.dumps(col_main)
        }
        r1 = requests.post("https://api.monday.com/v2", headers=headers, json={"query": query1, "variables": vars1})
        data1 = r1.json()
        if "errors" in data1:
            print(f"❌ Monday Main: {data1['errors']}")
        else:
            print(f"✅ Fast projects Estimate карточка создана: {data1}")

        # Contacts board
        # phone=Phone, email=Email, location__1=Location, text4=Notes
        col_contacts = {
            "phone": {"phone": client_phone.replace("(", "").replace(")", "").replace(" ", "").replace("-", ""), "countryShortName": "US"} if client_phone else {},
            "email": {"email": client_email, "text": client_email} if client_email else {},
            "text__1": location,
            "text4": job_type,
        }
        query2 = """mutation ($board: ID!, $name: String!, $cols: JSON!) {
            create_item(board_id: $board, item_name: $name, column_values: $cols) { id }
        }"""
        vars2 = {
            "board": MONDAY_CONTACTS_BOARD,
            "name": client_name,
            "cols": json.dumps(col_contacts)
        }
        r2 = requests.post("https://api.monday.com/v2", headers=headers, json={"query": query2, "variables": vars2})
        data2 = r2.json()
        if "errors" in data2:
            print(f"❌ Monday Contacts: {data2['errors']}")
        else:
            print(f"✅ Contacts карточка создана: {data2}")

        return True
    except Exception as e:
        print(f"❌ Ошибка: {e}")
        return False
