import os
import re
import base64
from anthropic import Anthropic
from dotenv import load_dotenv
from rag import search_prices, init_database
from integrations import (
    upload_to_dropbox,
    upload_photo_to_client_folder,
    send_estimate_email,
    create_monday_cards,
    find_client_in_monday,
    find_client_by_phone,
    get_dropbox_photos,
    update_monday_item_link
)

load_dotenv()

client = Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
conversations = {}
init_database()

SYSTEM_PROMPT_BASE = """You are AI Estimator Agent for Handy Eddy handyman business (Los Angeles / Pasadena, CA area).

OWNER: Adam (Handy Eddy) — you work FOR him, not for the client.
COMPANY: Qirim LLC | 1105 E Orange Grove Blvd, Pasadena, CA 91104 | adam@handyeddy.com | 323 763 8686

═══════════════════════════════════════
YOUR JOB — IN ORDER OF PRIORITY:
═══════════════════════════════════════

1. CREATE ESTIMATES — your primary function
2. SEND EMAILS — when Adam approves
3. ANSWER QUESTIONS — about jobs, materials, timelines
4. SEARCH PRICES — when asked

═══════════════════════════════════════
ESTIMATE FORMAT — ALWAYS EXACT:
═══════════════════════════════════════

CLIENT: [full name]
PHONE: [phone or N/A]
EMAIL: [email or N/A]
JOB: [job description]
LOCATION: [street address or city, state — NEVER put address in NOTES]
BREAKDOWN:
- [task description] (labor): $[amount]
- [materials description] (materials): $[amount]
SUBTOTAL MATERIALS: $[amount]
SUBTOTAL LABOR: $[amount]
TOTAL ESTIMATE: $[amount]
TIMELINE: [e.g. 1-2 days]
NOTES: [scope clarifications, assumptions, optional add-ons — NOT the address]

RULES:
- Use N/A for missing phone/email/location — never ask for them, generate estimate anyway
- LOCATION = physical job address (street or city/state). NOTES = work scope details only
- LA/Pasadena labor rates: handyman $75-150/hr, painter $85-100/hr, plumber $100-200/hr
- Always price for LA market — not national average
- If photos provided — analyze them, describe what you see, price accordingly
- If no photos — price based on description

═══════════════════════════════════════
MONDAY WORKFLOW:
═══════════════════════════════════════

When someone tags @AI_Assist in Monday:
- Read the card columns: client name, phone, email, location, job description
- Check if photos exist in Dropbox folder for this client
- Generate estimate immediately — do not ask for more info unless critical
- After estimate is generated: Excel created → uploaded to Dropbox → link updated in Monday card

When Adam writes "estimate approve send to client":
- Send the estimate via DocuSign to client email
- Confirm in Monday comment when sent

═══════════════════════════════════════
PRICING GUIDELINES — LA/PASADENA:
═══════════════════════════════════════

Doors: pre-hung interior $250-400 labor each | exterior $350-600 labor each
Painting interior: $85-95/hr labor | room $300-800 | whole house $4,000-10,000
Painting exterior: $2,000-8,000 depending on size
Plumbing: $100-200/hr | faucet $150-350 | toilet $200-450 | sink $200-500
Drywall repair: small hole $150-300 | water damage $400-1,200 | full room $1,200-3,500
Flooring: hardwood $8-15/sqft | vinyl plank $3-7/sqft | laminate $4-9/sqft
Tile: $10-30/sqft installed
Stucco repair: small $200-400 | large patch $500-1,500
Popcorn ceiling removal: $300-600 per room
Garage door: single $800-1,500 | double $1,500-3,500 installed
Handyman general: $75-150/hr | half day $300-550 | full day $500-900
Carpentry/trim: $85-175/hr
Electrical: $100-175/hr
Fence: wood $20-50/linear ft | block wall $50-100/linear ft
Pressure washing: house $300-800 | driveway $150-350
Window install: $400-1,000 per window
Cabinet install: $300-500 per cabinet

═══════════════════════════════════════
HOW TO THINK WHEN ESTIMATING — critical logic:
═══════════════════════════════════════

You are not a database — you are an experienced LA handyman with 15+ years of field experience. Before writing any estimate, think through these steps:

**Step 1 — Assess the REAL scope**
Price ranges in RAG data are WIDE on purpose (e.g. "bathroom remodel $4k-$32k"). Your job is to pick the RIGHT point inside that range based on what you actually see. Don't default to the middle — analyze the job:
- If photos/description show "everything is fine, just needs refresh" → lean to BOTTOM of range
- If "moderate wear, standard job" → MIDDLE of range
- If "major damage, old building, complex access, premium finishes" → TOP of range
- If description is vague → ask yourself "what's the MOST LIKELY reality" and pick that point

**Step 2 — Break down into specific tasks**
Never give a lump sum like "bathroom remodel $15,000". A real handyman breaks work into concrete tasks:
- Demolition/removal — how much material to tear out?
- Prep work — surface prep, masking, protection of other areas
- Core work — actual labor with hours × rate
- Materials — specific items with quantities
- Finishing — cleanup, touch-ups, haul-away
Each line in BREAKDOWN must be a SPECIFIC task someone could actually do, not a category.

**Step 3 — Consider LA/Pasadena context**
- Many Pasadena homes are 1920s-1960s — expect plaster walls (not drywall), lath backing, lead paint risk (pre-1978 homes)
- Stucco exteriors are common — stucco repair, stucco painting often needed
- Asbestos risk in old popcorn ceilings (pre-1980)
- Tile roofs common — different from asphalt shingle pricing
- Earthquake-related cracks are common — cosmetic vs structural matters
- Mediterranean/Spanish style homes may have unique architectural details
- If pre-1978 home and lead paint suspected — add $500-$800 for lead-safe practices

**Step 4 — Use actual hours × rate for labor**
Instead of "labor: $600" — think "this will take a 2-man crew 6 hours at $95/hr = $1,140". Round to nearest $25-50. Always show your reasoning implicitly in the breakdown:
- "Remove old fixture and patch drywall (2 hrs labor): $200"
- "Install new vanity and plumbing hookup (3 hrs): $350"

**Step 5 — Add professional realism**
Every estimate should include these (even if small):
- Protection/masking of nearby areas (for messy work)
- Cleanup and haul-away of debris
- Minor unforeseen (budget 10-15% or flag as "may discover during work")
- Permits if needed (for electrical/plumbing/structural changes in LA)

**Step 6 — Sanity check your total**
Before finalizing:
- Is this total reasonable for Adam's typical clients in LA/Pasadena?
- Would a client think "fair price" or "wtf too high/low"?
- If total feels off — recheck your breakdown, something is probably missing or inflated

**Remember:** RAG prices are REFERENCE ranges, not answers. Your job is to APPLY expert judgment to the specific case in front of you, using photos/description to narrow down where in the range this particular job lands.

**Step 7 — Consistency rule (CRITICAL)**
If the same photos/description are presented to you twice, your estimate MUST be within 10% of itself. To achieve this:
- Always start from measurement: count sq ft, count linear ft, count items
- Always use the SAME labor rate ($95/hr for 2-man crew standard)
- Never round to "nice" numbers — calculate from hours × rate precisely
- Never change your mind about scope between runs — if you see lead paint risk once, you ALWAYS see it
- If you previously estimated this job and you're being asked again — your new total should be IDENTICAL unless the scope explicitly changed

**Step 8 — Measurement discipline**
Before writing any dollar amount, ALWAYS state the measurement:
- "Exterior walls: ~1,400 sq ft (160 linear ft × 9 ft avg)"
- "Interior trim: ~180 linear ft"
- "Cabinet doors: 22 doors"
- "Tile area: 85 sq ft"
This measurement is the foundation. Dollar amounts flow from it. Same measurement = same dollars.

═══════════════════════════════════════
COMMUNICATION STYLE:
═══════════════════════════════════════

- Professional, concise, no fluff
- English only (unless client writes in another language)
- Never say you cannot access Dropbox, Monday, or email — system handles everything
- Never ask for info that is already in the Monday card columns
- Maximum response 3,500 tokens — never cut off mid-sentence

═══════════════════════════════════════
═══════════════════════════════════════

- Professional, concise, no fluff
- English only (unless client writes in another language)
- Never say you cannot access Dropbox, Monday, or email — system handles everything
- Never ask for info that is already in the Monday card columns
- Maximum response 3,500 tokens — never cut off mid-sentence

═══════════════════════════════════════
SYSTEM HANDLES AUTOMATICALLY:
═══════════════════════════════════════
Excel generation, Dropbox upload, Monday card update (link column), Email sending via SendGrid, Photo retrieval from Dropbox."""


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


def is_estimate_request(message):
    return any(x in message.lower() for x in [
        "estimate", "quote", "price", "cost", "how much",
        "make estimate", "create estimate", "prepare estimate",
        "эстимейт", "посчитай", "сколько стоит"
    ])


# ─────────────────────────────────────────
# PIPELINE — Telegram
# ─────────────────────────────────────────

async def run_pipeline(assistant_message, file_number="", known_phone=""):
    """
    Создаём Excel → загружаем в Dropbox → создаём карточки в Monday.
    file_number — если есть (из Monday карточки), используется для синхронизации.
    known_phone — если есть, ищем существующего клиента по телефону для реюза номера.
    """
    from excel_generator import generate_estimate_excel as generate_pdf
    data        = extract_estimate_data(assistant_message)
    client_name = data.get("client_name", "Client")
    job_type    = data.get("job_type", "Job")
    total       = data.get("total", "0")
    phone       = data.get("phone", "") or known_phone
    email       = data.get("email", "")
    location    = data.get("location", "")

    try:
        # Если телефон есть — проверим существующего клиента по нему
        if phone and not file_number:
            existing = find_client_by_phone(phone)
            if existing and existing.get("file_number"):
                file_number = existing["file_number"]
                print(f"Повторный клиент по телефону: reusing number {file_number}")

        pdf_path = generate_pdf(assistant_message, client_name, job_type)

        dropbox_url, used_number = upload_to_dropbox(
            pdf_path, client_name, job_type,
            file_number=file_number or None
        )

        monday_result = create_monday_cards(
            client_name, job_type, total,
            dropbox_url, email, phone, location,
            file_number=used_number
        )

        return (
            f"\n\n✅ Excel создан\n"
            f"📁 Dropbox: {dropbox_url}\n"
            f"🔢 #: {used_number}\n"
            f"📋 Monday: {monday_result}"
        )
    except Exception as e:
        return f"\n\n⚠️ Пайплайн ошибка: {str(e)}"


# ─────────────────────────────────────────
# PIPELINE — Monday
# ─────────────────────────────────────────

async def run_pipeline_monday(assistant_message, client_name, item_id=None, file_number=""):
    """
    Pipeline для Monday webhook. Использует file_number из карточки
    чтобы положить Excel в существующую папку Dropbox (ту же что и карточка).
    """
    from excel_generator import generate_estimate_excel as generate_pdf
    data     = extract_estimate_data(assistant_message)
    job_type = data.get("job_type", "Job")
    try:
        pdf_path = generate_pdf(assistant_message, client_name, job_type)
        dropbox_url, used_number = upload_to_dropbox(
            pdf_path, client_name, job_type,
            file_number=file_number or None
        )
        if item_id:
            update_monday_item_link(item_id, dropbox_url)
        return f'\n\n✅ Excel создан\n📁 <a href="{dropbox_url}">Открыть в Dropbox</a>'
    except Exception as e:
        return f"\n\n⚠️ Пайплайн ошибка: {str(e)}"


# ─────────────────────────────────────────
# EMAIL HANDLER
# ─────────────────────────────────────────

async def handle_email_command(message, context_dropbox_url="", context_job_type="", context_client_name=""):
    try:
        email_in_message = re.search(r"[\w.+-]+@[\w.-]+\.[a-zA-Z]{2,}", message)

        if email_in_message:
            direct_email = email_in_message.group(0)
            sent = send_estimate_email(
                client_email=direct_email,
                client_name=context_client_name or "Client",
                job_type=context_job_type or "Handyman Estimate",
                pdf_path=None,
                dropbox_url=context_dropbox_url or ""
            )
            if sent:
                return f"✅ Email отправлен на {direct_email}"
            else:
                return f"❌ Ошибка отправки на {direct_email}"

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
            lines = [f"🔍 Найдено {len(result)} клиента с именем '{client_name}'. Уточни:\n"]
            for c in result:
                location = c.get("location") or "локация не указана"
                job      = c.get("job_type") or "работа не указана"
                number   = c.get("file_number") or "?"
                lines.append(f"• {c['name']} [#{number}] — {location} | {job}")
            lines.append("\nНапиши полное имя с номером: Send email to Maria Gonzalez 26041126")
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
# MONDAY REQUEST
# ─────────────────────────────────────────

async def process_monday_request(item_id, client_name, job_type,
                                  location, phone, email,
                                  dropbox_url, message,
                                  file_number="", notes=""):
    try:
        user_key = f"monday_{item_id}"

        intent = detect_intent(message)
        need_estimate = is_estimate_request(message)

        # Для эстимейтов всегда начинаем с чистого листа — это критично для
        # стабильности цифр. Иначе предыдущий эстимейт в памяти влияет на новый.
        if need_estimate:
            conversations[user_key] = []
        elif user_key not in conversations:
            conversations[user_key] = []

        if intent == "email":
            return await handle_email_command(
                message,
                context_dropbox_url=dropbox_url,
                context_job_type=job_type,
                context_client_name=client_name
            )

        message_content = []

        context_header = (
            f"Client: {client_name}. "
            f"Phone: {phone or 'N/A'}. "
            f"Email: {email or 'N/A'}. "
            f"Job: {job_type}. "
            f"Location: {location or 'N/A'}. "
            f"Notes: {notes or 'N/A'}. "
            f"File #: {file_number or 'N/A'}.\n"
        )

        if need_estimate:
            # Ищем фото — по номеру если есть, иначе по имени
            photos = get_dropbox_photos(client_name, file_number=file_number or None)
            if photos:
                message_content.append({
                    "type": "text",
                    "text": (
                        context_header +
                        f"Request: {message}\n"
                        f"Analyze these {len(photos)} photos and create estimate. "
                        f"Use phone/email/location from above — do NOT ask for them again."
                    )
                })
                for photo in photos:
                    message_content.append({
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": "image/jpeg",
                            "data": photo["data"]
                        }
                    })
                print(f"Loaded {len(photos)} photos from Dropbox for {client_name}")
            else:
                message_content.append({
                    "type": "text",
                    "text": (
                        context_header +
                        f"Request: {message}\n"
                        f"No photos found in Dropbox. "
                        f"Create estimate based on job description. "
                        f"Use phone/email/location from above — do NOT ask for them again."
                    )
                })
        else:
            message_content.append({
                "type": "text",
                "text": context_header + f"Question: {message}"
            })

        conversations[user_key].append({"role": "user", "content": message_content})

        prices        = search_prices(message or job_type)
        system_prompt = SYSTEM_PROMPT_BASE + f"\n\nMarket prices from RAG: {prices}"

        kwargs = {
            "model":       "claude-opus-4-6",
            "max_tokens":  4000,
            "temperature": 0,
            "system":      system_prompt,
            "messages":    conversations[user_key]
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

        conversations[user_key].append({"role": "assistant", "content": assistant_message})

        if "TOTAL ESTIMATE:" in assistant_message and need_estimate:
            pipeline_result = await run_pipeline_monday(
                assistant_message,
                client_name=client_name,
                item_id=item_id,
                file_number=file_number
            )
            assistant_message += pipeline_result

        return assistant_message

    except Exception as e:
        import traceback
        traceback.print_exc()
        return f"❌ Ошибка: {str(e)}"


# ─────────────────────────────────────────
# TELEGRAM — MAIN PROCESS
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
            "model":       "claude-opus-4-6",
            "max_tokens":  4000,
            "temperature": 0,
            "system":      system_prompt,
            "messages":    conversations[user_id]
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
        import traceback
        traceback.print_exc()
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
            temperature=0,
            system=system_prompt,
            messages=conversations[user_id]
        )
        assistant_message = response.content[0].text
        conversations[user_id].append({"role": "assistant", "content": assistant_message})

        # Сохраняем само фото в папку клиента (если сумели вытащить имя)
        saved_number = ""
        try:
            data = extract_estimate_data(assistant_message)
            client_name_guess = data.get("client_name") or caption or ""
            phone_guess = data.get("phone", "")

            # Если есть телефон — пробуем найти существующего клиента для реюза номера
            existing_number = ""
            if phone_guess and phone_guess.lower() not in ["n/a", "na", ""]:
                existing = find_client_by_phone(phone_guess)
                if existing and existing.get("file_number"):
                    existing_number = existing["file_number"]

            if client_name_guess:
                saved_number = upload_photo_to_client_folder(
                    image_data, client_name_guess,
                    file_number=existing_number or None
                ) or ""
                print(f"Photo saved to Dropbox under number {saved_number}")
        except Exception as e:
            print(f"Photo save skipped: {e}")

        if "TOTAL ESTIMATE:" in assistant_message:
            pipeline_result = await run_pipeline(
                assistant_message,
                file_number=saved_number
            )
            assistant_message += pipeline_result

        return assistant_message

    except Exception as e:
        import traceback
        traceback.print_exc()
        return f"❌ Ошибка: {str(e)}"


async def process_dropbox_photos(user_id, client_name, caption="", file_number=""):
    try:
        # Ищем фото — по номеру если есть, иначе по имени
        photos = get_dropbox_photos(client_name, file_number=file_number or None)
        if not photos:
            return f"❌ Фото не найдены в Dropbox для: {client_name} {file_number}"

        if user_id not in conversations:
            conversations[user_id] = []

        if caption:
            instruction = (
                f"Analyze photos for {client_name} (file #{file_number or 'N/A'}). "
                f"Scope specified: {caption}. "
                f"Create estimate for ONLY what was specified."
            )
        else:
            instruction = (
                f"Analyze photos for {client_name} (file #{file_number or 'N/A'}). "
                f"Describe what work you see needed. "
                f"Ask: 'Is this the correct scope or do you want to adjust?' "
                f"Do NOT generate TOTAL ESTIMATE yet — wait for confirmation."
            )

        message_content = [{"type": "text", "text": instruction}]
        for photo in photos:
            message_content.append({
                "type": "image",
                "source": {"type": "base64", "media_type": "image/jpeg", "data": photo["data"]}
            })

        conversations[user_id].append({"role": "user", "content": message_content})

        prices        = search_prices(caption or "handyman work")
        system_prompt = SYSTEM_PROMPT_BASE + f"\n\nMarket prices from RAG: {prices}"

        response = client.messages.create(
            model="claude-opus-4-6",
            max_tokens=4000,
            temperature=0,
            system=system_prompt,
            messages=conversations[user_id]
        )
        assistant_message = response.content[0].text
        conversations[user_id].append({"role": "assistant", "content": assistant_message})

        if "TOTAL ESTIMATE:" in assistant_message and caption:
            pipeline_result = await run_pipeline(
                assistant_message,
                file_number=file_number
            )
            assistant_message += pipeline_result

        return assistant_message

    except Exception as e:
        import traceback
        traceback.print_exc()
        return f"❌ Ошибка: {str(e)}"


def clear_conversation(user_id):
    if user_id in conversations:
        del conversations[user_id]
    return "🔄 Разговор сброшен. Начнём заново!"
