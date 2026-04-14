import os
import re
import base64
from anthropic import Anthropic
from dotenv import load_dotenv
from rag import search_prices, init_database
from integrations import (
    upload_to_dropbox,
    send_estimate_email,
    create_monday_cards,
    find_client_in_monday
)

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
- [item materials]: $[price]
- [item labor]: $[price]
SUBTOTAL MATERIALS: $[amount]
SUBTOTAL LABOR: $[amount]
TOTAL ESTIMATE: $[amount]
TIMELINE: [days]
NOTES: [notes]

=== EDIT MODE ===
When asked to modify estimate (remove item, recalculate, change price):
- Update the breakdown
- Recalculate all totals
- Output full estimate again in EXACT format above

=== EMAIL MODE ===
When asked to send email — respond: "Sending estimate to [name]..."
System handles the actual sending automatically.

=== WEB SEARCH MODE ===
When asked to find prices or product info — use web_search tool.
Return concise result: product name, price, source.

=== GENERAL ===
- Be concise. No fluff.
- If asked anything work-related (materials, timelines, techniques) — answer directly.
- English only unless client writes in another language.
- Never say you cannot access Dropbox, Monday or email — system handles all of that automatically.

IMPORTANT: Keep all responses under 3500 tokens.
For estimates — be concise in NOTES and descriptions.
Never cut off mid-sentence. Always complete your response fully."""


# ─────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────

def extract_estimate_data(text):
    data = {}
    patterns = {
        "client_name": r"CLIENT:\s*(.+)",
        "job_type":    r"JOB:\s*(.+)",
        "location":    r"LOCATION:\s*(.+)",
        "total":       r"\*{0,2}TOTAL ESTIMATE:\*{0,2}\s*:?\s*\$?([\d,]+)",
        "phone":       r"PHONE:\s*(.+)",
        "email":       r"EMAIL:\s*(.+)",
    }
    for key, pattern in patterns.items():
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            data[key] = match.group(1).strip().lstrip("*").strip()
    return data


def detect_intent(message):
    msg = message.lower()
    if any(x in msg for x in [
        "send email", "отправь имейл", "отправь email",
        "email to", "имейл", "отправь письмо"
    ]):
        return "email"
    if any(x in msg for x in [
        "search", "find price", "how much is", "web search",
        "найди цену", "погугли", "поищи", "сколько стоит"
    ]):
        return "search"
    return "estimate"


# ─────────────────────────────────────────
# PIPELINE
# ─────────────────────────────────────────

async def run_pipeline(assistant_message, existing_client_folder=None):
    from excel_generator import generate_estimate_excel as generate_pdf
    data        = extract_estimate_data(assistant_message)
    client_name = data.get("client_name", "Client")
    job_type    = data.get("job_type", "Job")
    total       = data.get("total", "0")
    phone       = data.get("phone", "")
    email       = data.get("email", "")
    location    = data.get("location", "")
    try:
        pdf_path      = generate_pdf(assistant_message, client_name, job_type)
        folder_to_use = existing_client_folder or client_name
        dropbox_url   = upload_to_dropbox(
            pdf_path, client_name, job_type,
            existing_folder=folder_to_use
        )
        monday_result = create_monday_cards(
            client_name, job_type, total,
            dropbox_url, email, phone, location
        )
        return (
            f"\n\n✅ Excel создан\n"
            f"📁 Dropbox: {dropbox_url}\n"
            f"📋 Monday: {monday_result}"
        )
    except Exception as e:
        return f"\n\n⚠️ Пайплайн ошибка: {str(e)}"


# ─────────────────────────────────────────
# EMAIL HANDLER
# ─────────────────────────────────────────

async def handle_email_command(message):
    try:
        patterns = [
            r"send email(?:\s+to)?\s+(.+)",
            r"отправь(?:\s+имейл|\s+email|\s+письмо)(?:\s+to|\s+клиенту)?\s+(.+)",
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

        result = find_client_in_monday(client_name)

        if result is None:
            return f"❌ Клиент '{client_name}' не найден в Monday."

        if isinstance(result, list):
            lines = [f"🔍 Найдено {len(result)} клиента с именем '{client_name}'. Уточни номер карточки:\n"]
            for c in result:
                location = c.get("location") or "локация не указана"
                job      = c.get("job_type") or "работа не указана"
                lines.append(f"• {c['name']} — {location} | {job}")
            lines.append("\nНапиши полное имя с номером, например: Send email to Maria Gonzalez 260414")
            return "\n".join(lines)

        client_email = result.get("email", "")
        dropbox_url  = result.get("dropbox_url", "")
        job_type     = result.get("job_type", "Handyman Work")
        name         = result.get("name", client_name)

        if not client_email or client_email.lower() in ["n/a", "na", ""]:
            return f"❌ У клиента {name} нет email в системе."

        sent = send_estimate_email(
            client_email=client_email,
            client_name=name,
            job_type=job_type,
            pdf_path=None,
            dropbox_url=dropbox_url
        )

        if sent:
            return f"✅ Email отправлен {name} на {client_email}\n📁 {dropbox_url}"
        else:
            return f"❌ Ошибка отправки для {name}"

    except Exception as e:
        return f"❌ Email error: {str(e)}"


# ─────────────────────────────────────────
# MAIN PROCESS
# ─────────────────────────────────────────

async def process_request(user_id, user_message):
    try:
        if user_id not in conversations:
            conversations[user_id] = []

        intent = detect_intent(user_message)

        if intent == "email":
            return await handle_email_command(user_message)

        conversations[user_id].append({"role": "user", "content": user_message})

        prices        = search_prices(user_message)
        system_prompt = SYSTEM_PROMPT_BASE + f"\n\nMarket prices from RAG: {prices}"

        kwargs = {
            "model":      "claude-opus-4-6",
            "max_tokens": 4000,
            "system":     system_prompt,
            "messages":   conversations[user_id]
        }

        if intent == "search":
            kwargs["tools"] = [{"type": "web_search_20250305", "name": "web_search"}]

        response = client.messages.create(**kwargs)

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

        prices        = search_prices(caption or "general handyman work")
        system_prompt = SYSTEM_PROMPT_BASE + f"\n\nMarket prices from RAG: {prices}"

        response = client.messages.create(
            model="claude-opus-4-6",
            max_tokens=4000,
            system=system_prompt,
            messages=conversations[user_id]
        )
        assistant_message = response.content[0].text
        conversations[user_id].append({"role": "assistant", "content": assistant_message})

        if "TOTAL ESTIMATE:" in assistant_message:
            data        = extract_estimate_data(assistant_message)
            client_name = data.get("client_name", "Client")
            pipeline_result = await run_pipeline(
                assistant_message,
                existing_client_folder=client_name
            )
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

        prices        = search_prices("exterior painting handyman work")
        system_prompt = SYSTEM_PROMPT_BASE + f"\n\nMarket prices from RAG: {prices}"

        response = client.messages.create(
            model="claude-opus-4-6",
            max_tokens=4000,
            system=system_prompt,
            messages=conversations[user_id]
        )
        assistant_message = response.content[0].text
        conversations[user_id].append({"role": "assistant", "content": assistant_message})

        if "TOTAL ESTIMATE:" in assistant_message:
            pipeline_result = await run_pipeline(
                assistant_message,
                existing_client_folder=client_name
            )
            assistant_message += pipeline_result

        return assistant_message

    except Exception as e:
        return f"❌ Ошибка: {str(e)}"


def clear_conversation(user_id):
    if user_id in conversations:
        del conversations[user_id]
    return "🔄 Разговор сброшен. Начнём заново!"
