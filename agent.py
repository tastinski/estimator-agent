import os
import re
import base64
from anthropic import Anthropic
from dotenv import load_dotenv
from rag import search_prices, init_database
from integrations import upload_to_dropbox, send_estimate_email, create_monday_cards, find_client_in_monday

load_dotenv()

client = Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
conversations = {}
init_database()

SYSTEM_PROMPT_BASE = """You are AI Assistant for Handy Eddy handyman business (USA).

YOUR PRIMARY JOB: Create professional estimates for handyman work.
SECONDARY: Help with any work-related tasks — send emails, search prices, answer questions about jobs.

=== ESTIMATE MODE ===
When client describes work → collect info → generate estimate in EXACT format:
CLIENT: [full name]
PHONE: [phone]
EMAIL: [email or N/A]
JOB: [job type]
LOCATION: [city, state]
BREAKDOWN:
- [item]: $[price]
SUBTOTAL MATERIALS: $[amount]
SUBTOTAL LABOR: $[amount]
TOTAL ESTIMATE: $[amount]
TIMELINE: [days]
NOTES: [notes]

=== EMAIL MODE ===
When asked to send email to a client:
- Confirm: "Sending estimate email to [name] at [email]..."
- System handles the actual sending automatically.
- Format: EMAIL_COMMAND: [client name]

=== WEB SEARCH MODE ===
When asked to find prices or product info — use web search tool.
Return concise results with source and price.

=== EDIT MODE ===
When asked to modify estimate (remove item, recalculate, change price):
- Update the breakdown
- Recalculate totals
- Output full estimate again in EXACT format

Always be concise. No fluff. English only unless client writes in another language."""


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


def detect_intent(message):
    msg = message.lower()
    if any(x in msg for x in ["send email", "отправь имейл", "отправь email", "email to", "имейл"]):
        return "email"
    if any(x in msg for x in ["search", "find price", "how much", "web search", "найди цену", "погугли", "поищи"]):
        return "search"
    return "estimate"


async def handle_email_command(user_id, message):
    """Parse client name and send estimate email"""
    try:
        from integrations import find_client_in_monday, send_estimate_email

        # Extract name - everything after trigger words
        patterns = [
            r"send email(?:\s+to)?\s+(.+)",
            r"отправь(?:\s+имейл|\s+email)(?:\s+to|\s+клиенту)?\s+(.+)",
            r"email\s+to\s+(.+)",
        ]
        client_name = None
        for pattern in patterns:
            match = re.search(pattern, message, re.IGNORECASE)
            if match:
                client_name = match.group(1).strip(" .,?!")
                break

        if not client_name:
            return "❌ Не понял имя клиента. Напиши: 'Send email to Maria Gonzalez'"

        # Find client in Monday
        client_data = find_client_in_monday(client_name)
        if not client_data:
            return f"❌ Клиент '{client_name}' не найден в Monday."

        client_email = client_data.get("email")
        dropbox_url = client_data.get("dropbox_url")
        job_type = client_data.get("job_type", "Handyman Work")
        name = client_data.get("name", client_name)

        if not client_email or client_email in ["N/A", "n/a", ""]:
            return f"❌ У клиента {name} нет email в системе."

        # Build simple estimate text from Monday data
        estimate_text = f"Estimate for {name} — {job_type}"

        # We need a local xlsx path — for now send without attachment, just Dropbox link
        result = send_estimate_email(
            client_email=client_email,
            client_name=name,
            job_type=job_type,
            estimate_text=estimate_text,
            pdf_path=None,
            dropbox_url=dropbox_url
        )

        if result:
            return f"✅ Email отправлен {name} на {client_email}\n📁 Dropbox: {dropbox_url}"
        else:
            return f"❌ Ошибка отправки email для {name}"

    except Exception as e:
        return f"❌ Email error: {str(e)}"


async def run_pipeline(assistant_message, existing_client_folder=None):
    from excel_generator import generate_estimate_excel as generate_pdf
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

        # Intent detection
        intent = detect_intent(user_message)

        # EMAIL command — handle separately
        if intent == "email":
            return await handle_email_command(user_id, user_message)

        conversations[user_id].append({"role": "user", "content": user_message})
        prices = search_prices(user_message)

        system_prompt = SYSTEM_PROMPT_BASE + f"\n\nMarket prices from RAG: {prices}"

        # WEB SEARCH — add tool if search intent
        tools = []
        if intent == "search":
            tools = [{"type": "web_search_20250305", "name": "web_search"}]

        kwargs = {
            "model": "claude-opus-4-6",
            "max_tokens": 2000,
            "system": system_prompt,
            "messages": conversations[user_id]
        }
        if tools:
            kwargs["tools"] = tools

        response = client.messages.create(**kwargs)

        # Extract text from response (web search returns mixed content blocks)
        assistant_message = ""
        for block in response.content:
            if hasattr(block, "text"):
                assistant_message += block.text

        if not assistant_message:
            assistant_message = "Готово."

        conversations[user_id].append({"role": "assistant", "content": assistant_message})

        if "TOTAL ESTIMATE:" in assistant_message:
            pipeline_result = await run_pipeline(assistant_message)
            assistant_message += pipeline_result

        return assistant_message

    except Exception as e:
        return f"❌ Ошибка: {str(e)}"


async def process_photo(user_id, image_data, caption=""):
    try:
        from excel_generator import generate_estimate_excel as generate_pdf
        image_b64 = base64.standard_b64encode(image_data).decode("utf-8")
        if user_id not in conversations:
            conversations[user_id] = []

        message_content = []
        if caption:
            message_content.append({"type": "text", "text": caption})
        message_content.append({
            "type": "image",
            "source": {"type": "base64", "media_type": "image/jpeg", "data": image_b64}
        })
        conversations[user_id].append({"role": "user", "content": message_content})

        prices = search_prices(caption or "general handyman work")
        system_prompt = SYSTEM_PROMPT_BASE + f"\n\nMarket prices from RAG: {prices}"

        response = client.messages.create(
            model="claude-opus-4-6",
            max_tokens=1500,
            system=system_prompt,
            messages=conversations[user_id]
        )
        assistant_message = response.content[0].text
        conversations[user_id].append({"role": "assistant", "content": assistant_message})

        if "TOTAL ESTIMATE:" in assistant_message:
            data = extract_estimate_data(assistant_message)
            client_name = data.get("client_name", "Client")
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

        message_content = [
            {"type": "text", "text": f"Analyze all these photos for {client_name} and create a detailed estimate."}
        ]
        for photo in photos:
            message_content.append({
                "type": "image",
                "source": {"type": "base64", "media_type": "image/jpeg", "data": photo["data"]}
            })

        conversations[user_id].append({"role": "user", "content": message_content})
        prices = search_prices("exterior painting handyman work")
        system_prompt = SYSTEM_PROMPT_BASE + f"\n\nMarket prices from RAG: {prices}"

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
