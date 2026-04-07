import os
import re
from anthropic import Anthropic
from dotenv import load_dotenv
from rag import search_prices, init_database
from integrations import upload_to_dropbox, send_estimate_email, create_monday_cards
from pdf_generator import generate_estimate_pdf as generate_pdf

load_dotenv()

client = Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
conversations = {}
init_database()

def extract_estimate_data(text: str) -> dict:
    """Парсит текст estimate и достаёт данные"""
    data = {}
    patterns = {
        "client_name": r"CLIENT:\s*(.+)",
        "job_type": r"JOB:\s*(.+)",
        "location": r"LOCATION:\s*(.+)",
        "total": r"TOTAL ESTIMATE:\s*\$?([\d,]+)",
    }
    for key, pattern in patterns.items():
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            data[key] = match.group(1).strip()
    return data

async def process_request(user_id: int, user_message: str) -> str:
    try:
        if user_id not in conversations:
            conversations[user_id] = []

        conversations[user_id].append({
            "role": "user",
            "content": user_message
        })

        prices = search_prices(user_message)

        system_prompt = f"""You are a professional estimator for a handyman business in the USA.
Your job is to create ACCURATE estimates. To do this:
STEP 1 - If the request is vague, ask 2-3 specific clarifying questions:
- Size of the area (sq ft)?
- Scope of work (partial or full)?
- Materials preference (budget / mid-range / premium)?
- Any special conditions (floor, access, moving plumbing)?
STEP 2 - Once you have enough details, create a precise estimate.
Use these real market prices:
{prices}
When ready to estimate, use this format EXACTLY:
CLIENT: [full name]
JOB: [job type]
LOCATION: [city, state]
BREAKDOWN:
- [item 1]: $[price]
- [item 2]: $[price]
TOTAL ESTIMATE: $[amount]
TIMELINE: [days]
NOTES: [important notes]
IMPORTANT:
- If not enough info → ask questions first
- If enough info → give estimate directly
- Be specific, not wide ranges
- Always respond in English"""

        response = client.messages.create(
            model="claude-opus-4-6",
            max_tokens=1000,
            system=system_prompt,
            messages=conversations[user_id]
        )

        assistant_message = response.content[0].text

        conversations[user_id].append({
            "role": "assistant",
            "content": assistant_message
        })

        # Если estimate готов — запускаем пайплайн
        if "TOTAL ESTIMATE:" in assistant_message:
            data = extract_estimate_data(assistant_message)
            client_name = data.get("client_name", "Client")
            job_type = data.get("job_type", "Job")
            total = data.get("total", "0")

            try:
                pdf_path = generate_pdf(assistant_message, client_name, job_type)
                dropbox_url = upload_to_dropbox(pdf_path, client_name, job_type)
                create_monday_cards(client_name, job_type, total, dropbox_url)
                assistant_message += f"\n\n✅ PDF создан\n📁 Dropbox: {dropbox_url}\n📋 Monday карточка создана"
            except Exception as e:
                assistant_message += f"\n\n⚠️ Пайплайн ошибка: {str(e)}"

        return assistant_message

    except Exception as e:
        return f"❌ Ошибка: {str(e)}"

def clear_conversation(user_id: int):
    if user_id in conversations:
        del conversations[user_id]
    return "🔄 Разговор сброшен. Начнём заново!"

async def process_photo(user_id: int, image_data: bytes, caption: str = "") -> str:
    try:
        import base64
        image_b64 = base64.standard_b64encode(image_data).decode("utf-8")
        
        if user_id not in conversations:
            conversations[user_id] = []

        message_content = []
        if caption:
            message_content.append({"type": "text", "text": caption})
        message_content.append({
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": "image/jpeg",
                "data": image_b64
            }
        })

        conversations[user_id].append({
            "role": "user",
            "content": message_content
        })

        prices = search_prices(caption or "general handyman work")

        system_prompt = f"""You are a professional estimator for a handyman business in the USA.
The client sent a photo of the work area. Analyze it carefully and:
1. Describe what you see (condition, size estimate, scope of work)
2. Ask 1-2 clarifying questions if needed
3. If enough info - provide estimate using format:
CLIENT: [name]
JOB: [job type]
LOCATION: [location]
BREAKDOWN:
- [item]: $[price]
TOTAL ESTIMATE: $[amount]
TIMELINE: [days]
NOTES: [notes]

Market prices reference:
{prices}"""

        response = client.messages.create(
            model="claude-opus-4-6",
            max_tokens=1000,
            system=system_prompt,
            messages=conversations[user_id]
        )

        assistant_message = response.content[0].text
        conversations[user_id].append({
            "role": "assistant",
            "content": assistant_message
        })

        if "TOTAL ESTIMATE:" in assistant_message:
            data = extract_estimate_data(assistant_message)
            client_name = data.get("client_name", "Client")
            job_type = data.get("job_type", "Job")
            total = data.get("total", "0")
            try:
                pdf_path = generate_pdf(assistant_message, client_name, job_type)
                dropbox_url = upload_to_dropbox(pdf_path, client_name, job_type)
                create_monday_cards(client_name, job_type, total, dropbox_url)
                assistant_message += f"\n\n✅ PDF создан\n📁 Dropbox: {dropbox_url}\n📋 Monday карточка создана"
            except Exception as e:
                assistant_message += f"\n\n⚠️ Пайплайн ошибка: {str(e)}"

        return assistant_message

    except Exception as e:
        return f"❌ Ошибка: {str(e)}"

async def process_photo(user_id: int, image_data: bytes, caption: str = "") -> str:
    try:
        import base64
        image_b64 = base64.standard_b64encode(image_data).decode("utf-8")
        
        if user_id not in conversations:
            conversations[user_id] = []

        message_content = []
        if caption:
            message_content.append({"type": "text", "text": caption})
        message_content.append({
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": "image/jpeg",
                "data": image_b64
            }
        })

        conversations[user_id].append({
            "role": "user",
            "content": message_content
        })

        prices = search_prices(caption or "general handyman work")

        system_prompt = f"""You are a professional estimator for a handyman business in the USA.
The client sent a photo of the work area. Analyze it carefully and:
1. Describe what you see (condition, size estimate, scope of work)
2. Ask 1-2 clarifying questions if needed
3. If enough info - provide estimate using format:
CLIENT: [name]
JOB: [job type]
LOCATION: [location]
BREAKDOWN:
- [item]: $[price]
TOTAL ESTIMATE: $[amount]
TIMELINE: [days]
NOTES: [notes]

Market prices reference:
{prices}"""

        response = client.messages.create(
            model="claude-opus-4-6",
            max_tokens=1000,
            system=system_prompt,
            messages=conversations[user_id]
        )

        assistant_message = response.content[0].text
        conversations[user_id].append({
            "role": "assistant",
            "content": assistant_message
        })

        if "TOTAL ESTIMATE:" in assistant_message:
            data = extract_estimate_data(assistant_message)
            client_name = data.get("client_name", "Client")
            job_type = data.get("job_type", "Job")
            total = data.get("total", "0")
            try:
                pdf_path = generate_pdf(assistant_message, client_name, job_type)
                dropbox_url = upload_to_dropbox(pdf_path, client_name, job_type)
                create_monday_cards(client_name, job_type, total, dropbox_url)
                assistant_message += f"\n\n✅ PDF создан\n📁 Dropbox: {dropbox_url}\n📋 Monday карточка создана"
            except Exception as e:
                assistant_message += f"\n\n⚠️ Пайплайн ошибка: {str(e)}"

        return assistant_message

    except Exception as e:
        return f"❌ Ошибка: {str(e)}"
