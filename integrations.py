import os
import requests
import dropbox
import sendgrid
import base64
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

def upload_to_dropbox(pdf_path: str, client_name: str, job_type: str) -> str:
    try:
        dbx = dropbox.Dropbox(DROPBOX_TOKEN)
        date_str = datetime.now().strftime("%Y-%m-%d")
        client_folder = client_name.replace(' ', '_')
        job_folder = job_type.replace(' ', '_')[:30]
        dropbox_folder = f"/Estimator/{client_folder}/{job_folder}_{date_str}"
        filename = os.path.basename(pdf_path)
        dropbox_path = f"{dropbox_folder}/{filename}"
        with open(pdf_path, 'rb') as f:
            dbx.files_upload(f.read(), dropbox_path, mode=WriteMode.overwrite)
        print(f"✅ PDF загружен: {dropbox_path}")
        try:
            link = dbx.sharing_create_shared_link(dropbox_path)
            shared_url = link.url
        except Exception:
            links = dbx.sharing_list_shared_links(path=dropbox_path)
            shared_url = links.links[0].url if links.links else dropbox_path
        print(f"🔗 Ссылка: {shared_url}")
        return shared_url
    except ApiError as e:
        print(f"❌ Dropbox ошибка: {e}")
        return ""
    except Exception as e:
        print(f"❌ Ошибка: {e}")
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
            html_content=f"""
<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
    <div style="background: #1a1a2e; padding: 30px; text-align: center;">
        <h1 style="color: white; margin: 0;">Your Estimate is Ready</h1>
    </div>
    <div style="padding: 30px;">
        <p>Hi <b>{client_name}</b>,</p>
        <p>Thank you for reaching out! Please find your estimate attached.</p>
        <div style="background: #f8f9fa; padding: 20px; border-radius: 8px; margin: 20px 0;">
            <h3 style="margin: 0 0 10px 0;">Job: {job_type}</h3>
            <p style="margin: 0; color: #666;">Full breakdown is attached as PDF</p>
        </div>
        <p>📁 <a href="{dropbox_url}">View estimate in Dropbox</a></p>
        <p>Please don't hesitate to reach out if you have any questions.</p>
        <p>Best regards,<br><b>Edem</b></p>
    </div>
    <div style="background: #f8f9fa; padding: 15px; text-align: center; color: #999; font-size: 12px;">
        This estimate is valid for 30 days
    </div>
</div>"""
        )
        attachment = Attachment(
            FileContent(pdf_data),
            FileName(f"estimate_{client_name.replace(' ', '_')}.pdf"),
            FileType("application/pdf"),
            Disposition("attachment")
        )
        message.attachment = attachment
        sg.send(message)
        print(f"✅ Email отправлен: {client_email}")
        return True
    except Exception as e:
        print(f"❌ Email ошибка: {e}")
        return False

def create_monday_cards(client_name: str, job_type: str, estimate_total: str, dropbox_url: str, client_email: str = "") -> bool:
    try:
        headers = {
            "Authorization": MONDAY_API_KEY,
            "Content-Type": "application/json"
        }

        column_values = {
            "label__1": {"label": "text"},
            "text6__1": job_type,
            "text__1": dropbox_url,
            "email__1": {"email": client_email, "text": client_email}
        }

        import json
        query1 = """
        mutation {
            create_item (
                board_id: 18406712174,
                group_id: "group_mm0ea2gw",
                item_name: "%s",
                column_values: "%s"
            ) { id }
        }
        """ % (client_name, json.dumps(json.dumps(column_values))[1:-1])

        r1 = requests.post("https://api.monday.com/v2", headers=headers, json={"query": query1})
        data1 = r1.json()
        if "errors" not in data1:
            print(f"✅ Monday карточка создана!")
        else:
            print(f"❌ Monday ошибка: {data1}")

        return True

    except Exception as e:
        print(f"❌ Ошибка: {e}")
        return False
