import os
import re
import base64
from anthropic import Anthropic
from dotenv import load_dotenv
from rag import search_prices, init_database
from integrations import upload_to_dropbox, send_estimate_email, create_monday_cards
from excel_generator import generate_estimate_excel as generate_pdf

load_dotenv()

client = Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
conversations = {}
init_database()

def extract_estimate_data(text):
    data = {}
    patterns = {
        "client_name": r"CLIENT:\s*(.+)",
        "job_type": r"JOB:\s*(.+)",
        "location": r"LOCATION:\s*(.+)",
        "total": r"\*{0,2}TOTAL ESTIMATE:\*{0,2}\s*:?\s*\$?([\d,]+)",
        "phone": r"PHONE:\s*(.+)",
        "email": r"EMAIL:\s*(.+)",
    }
    for key, pattern in patterns.items():
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            data[key] = match.group(1).strip().lstrip("*").strip()
    return data

async def run_pipeline(assistant_message, existing_client_folder=None):
    data = extract_estimate_data(assistant_message)
    client_name = data.get("client_name", "Client")
    job_type = data.get("job_type", "Job")
    total = data.get("total", "0")
    phone = data.get("phone", "")
    email = data.get("email", "")
    location = data.get("location", "")
    try:
        pdf_path = generate_pdf(assistant_message, client_name, job_type)
        dropbox_url = upload_to_dropbox(pdf_path, client_name, job_type, existing_folder=existing_client_folder)
        monday_result = create_monday_cards(client_name, job_type, total, dropbox_url, email, phone, location)
        return f"\n\n✅ Excel создан\n📁 Dropbox: {dropbox_url}\n📋 Monday: {monday_result}"
    except Exception as e:
        return f"\n\n⚠️ Пайплайн ошибка: {str(e)}"

async def process_request(user_id, user_message):
    try:
        if user_id not in conversations:
            conversations[user_id] = []
        conversations[user_id].append({"role": "user", "content": user_message})
        prices = search_prices(user_message)
        system_prompt = f"""You are a professional estimator for Handy Eddy handyman business USA.
System automatically handles Dropbox, Monday.com, email. Never say you cannot access them.
Collect: client name, phone, email, job type, location, scope, materials.
When ready use EXACT format:
CLIENT: [full name]
PHONE: [phone]
EMAIL: [email or N/A]
JOB: [job type]
LOCATION: [city, state]
BREAKDOWN:
- [item]: $[price]
TOTAL ESTIMATE: $[amount]
TIMELINE: [days]
NOTES: [notes]
Market prices: {prices}"""
        response = client.messages.create(
            model="claude-opus-4-6",
            max_tokens=1500,
            system=system_prompt,
            messages=conversations[user_id]
        )
        assistant_message = response.content[0].text
        conversations[user_id].append({"role": "assistant", "content": assistant_message})
        if "TOTAL ESTIMATE:" in assistant_message:
            pipeline_result = await run_pipeline(assistant_message)
            assistant_message += pipeline_result
        return assistant_message
    except Exception as e:
        return f"❌ Ошибка: {str(e)}"

async def process_photo(user_id, image_data, caption=""):
    try:
        image_b64 = base64.standard_b64encode(image_data).decode("utf-8")
        if user_id not in conversations:
            conversations[user_id] = []
        message_content = []
        if caption:
            message_content.append({"type": "text", "text": caption})
        message_content.append({"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": image_b64}})
        conversations[user_id].append({"role": "user", "content": message_content})
        prices = search_prices(caption or "general handyman work")
        system_prompt = f"""You are a professional estimator for Handy Eddy handyman business USA.
Analyze the photo. Use EXACT format:
CLIENT: [name]
PHONE: [phone]
EMAIL: [email or N/A]
JOB: [job type]
LOCATION: [city, state]
BREAKDOWN:
- [item]: $[price]
TOTAL ESTIMATE: $[amount]
TIMELINE: [days]
NOTES: [notes]
Market prices: {prices}"""
        response = client.messages.create(
            model="claude-opus-4-6",
            max_tokens=1500,
            system=system_prompt,
            messages=conversations[user_id]
        )
        assistant_message = response.content[0].text
        conversations[user_id].append({"role": "assistant", "content": assistant_message})
        if "TOTAL ESTIMATE:" in assistant_message:
            pipeline_result = await run_pipeline(assistant_message, existing_client_folder=client_name)
            assistant_message += pipeline_result
        return assistant_message
    except Exception as e:
        return f"❌ Ошибка: {str(e)}"

async def process_dropbox_photos(user_id, client_name, caption=""):
    try:
        from integrations import get_dropbox_photos
        photos = get_dropbox_photos(client_name)
        if not photos:
            return f"❌ Фото не найдены в Dropbox для клиента: {client_name}"

        if user_id not in conversations:
            conversations[user_id] = []

        message_content = []
        message_content.append({"type": "text", "text": f"Analyze all these photos for {client_name} and create a detailed estimate."})

        for photo in photos:
            message_content.append({
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": "image/jpeg",
                    "data": photo["data"]
                }
            })

        conversations[user_id].append({"role": "user", "content": message_content})
        prices = search_prices("exterior painting handyman work")

        system_prompt = f"""You are a professional estimator for Handy Eddy handyman business USA.
Analyze ALL photos carefully. Keep response under 3000 tokens - be concise.
Use EXACT plain text format - NO markdown tables, NO ### headers:

CLIENT: {client_name}
PHONE: N/A
EMAIL: N/A
JOB: [job type based on photos]
LOCATION: [location if visible, otherwise Unknown]
BREAKDOWN:
- [Item materials description]: $[price]
- [Item labor description]: $[price]
SUBTOTAL MATERIALS: $[amount]
SUBTOTAL LABOR: $[amount]
TOTAL ESTIMATE: $[amount]
TIMELINE: [days]
NOTES: [brief notes based on photos]
Market prices: {prices}"""

        response = client.messages.create(
            model="claude-opus-4-6",
            max_tokens=3000,
            system=system_prompt,
            messages=conversations[user_id]
        )
        assistant_message = response.content[0].text
        conversations[user_id].append({"role": "assistant", "content": assistant_message})

        if "TOTAL ESTIMATE:" in assistant_message:
            pipeline_result = await run_pipeline(assistant_message)
            assistant_message += pipeline_result

        return assistant_message
    except Exception as e:
        return f"❌ Ошибка: {str(e)}"

def clear_conversation(user_id):
    if user_id in conversations:
        del conversations[user_id]
    return "🔄 Разговор сброшен. Начнём заново!"
