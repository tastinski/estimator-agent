import os
import httpx
import json
import asyncio
import threading
import traceback

from dotenv import load_dotenv
from flask import Flask, request, Response
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes

load_dotenv()

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
WEBHOOK_URL    = "https://estimator-agent-production.up.railway.app/telegram"

app       = Flask(__name__)
bot_app   = None
bot_loop  = None
bot_ready = threading.Event()


# ─────────────────────────────────────────
# TELEGRAM HANDLERS
# ─────────────────────────────────────────

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Привет! Я Estimator Agent.\n\n"
        "Опиши работу или скинь фото.\n\n"
        "/new — новый эстимейт\n"
        "/dropbox [номер имя] [scope] — анализ фото из Dropbox\n\n"
        "Примеры:\n"
        "/dropbox 26041314 Roseanne\n"
        "/dropbox 26041314 Roseanne pressure wash only"
    )

async def new_estimate(update: Update, context: ContextTypes.DEFAULT_TYPE):
    from agent import clear_conversation
    user_id = update.message.from_user.id
    result  = clear_conversation(user_id)
    await update.message.reply_text(result)

async def handle_dropbox(update: Update, context: ContextTypes.DEFAULT_TYPE):
    from agent import process_dropbox_photos
    user_id = update.message.from_user.id
    args    = context.args if context.args else []

    if not args:
        await update.message.reply_text(
            "Напиши имя клиента:\n"
            "/dropbox 26041314 Roseanne\n"
            "Или с уточнением: /dropbox 26041314 Roseanne pressure wash only"
        )
        return

    # Первый аргумент может быть номером (8 цифр) или именем
    first_arg = args[0]
    if first_arg.isdigit() and len(first_arg) == 8:
        file_number = first_arg
        client_name = " ".join(args[1:]) if len(args) > 1 else ""
        caption     = ""
        # Если после имени ещё слова — считаем их caption-ом
        # Формат: /dropbox 26041314 Roseanne pressure wash only
        # Тут непросто понять где кончается имя — берём первое слово как имя, остальное caption
        if len(args) >= 3:
            client_name = args[1]
            caption     = " ".join(args[2:])
    else:
        file_number = ""
        client_name = args[0]
        caption     = " ".join(args[1:]) if len(args) > 1 else ""

    await update.message.reply_text(
        f"📸 Ищу фото в Dropbox для {client_name or file_number}... Это займёт 1-2 минуты."
    )
    try:
        result     = await process_dropbox_photos(user_id, client_name, caption=caption, file_number=file_number)
        chunk_size = 4000
        for i in range(0, len(result), chunk_size):
            await update.message.reply_text(result[i:i + chunk_size])
    except Exception as e:
        print(f"handle_dropbox error: {e}")
        await update.message.reply_text(f"❌ Ошибка: {str(e)}")

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    from agent import process_request
    user_id      = update.message.from_user.id
    user_message = update.message.text
    try:
        await update.message.reply_text("Обрабатываю...")
        result = await process_request(user_id, user_message)
        await update.message.reply_text(result)
    except Exception as e:
        print(f"handle_message error: {e}")
        await update.message.reply_text(f"❌ Ошибка: {str(e)}")

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    from agent import process_photo
    user_id = update.message.from_user.id
    caption = update.message.caption or ""
    try:
        await update.message.reply_text("Фото получено, анализирую...")
        photo = update.message.photo[-1]
        file  = await context.bot.get_file(photo.file_id)
        async with httpx.AsyncClient() as client:
            response   = await client.get(file.file_path)
            image_data = response.content
        result = await process_photo(user_id, image_data, caption)
        await update.message.reply_text(result)
    except Exception as e:
        print(f"handle_photo error: {e}")
        await update.message.reply_text(f"❌ Ошибка анализа фото: {str(e)}")


# ─────────────────────────────────────────
# FLASK ROUTES
# ─────────────────────────────────────────

@app.route("/health", methods=["GET"])
def health():
    return "OK"

@app.route("/telegram", methods=["POST"])
def telegram_webhook():
    global bot_app, bot_loop
    if bot_app and bot_loop:
        data   = request.get_json(force=True)
        update = Update.de_json(data, bot_app.bot)
        asyncio.run_coroutine_threadsafe(bot_app.process_update(update), bot_loop)
    return Response(status=200)

@app.route("/sms", methods=["POST"])
def sms_reply():
    from twilio.twiml.messaging_response import MessagingResponse
    from agent import process_request
    phone = request.form.get("From", "")
    body  = request.form.get("Body", "").strip()
    try:
        future = asyncio.run_coroutine_threadsafe(
            process_request(phone, body), bot_loop
        )
        result = future.result(timeout=60)
    except Exception as e:
        result = f"Ошибка: {str(e)}"
    resp = MessagingResponse()
    if len(result) > 1600:
        result = result[:1597] + "..."
    resp.message(result)
    return str(resp)

@app.route("/monday_webhook", methods=["POST"])
def monday_webhook():
    try:
        data = request.get_json(force=True)

        if "challenge" in data:
            print(f"Monday CHALLENGE: {data['challenge']}", flush=True)
            return {"challenge": data["challenge"]}, 200

        # ВСЕГДА логируем всё что прилетает — видим в Railway логах
        print(f"Monday event received: {json.dumps(data)[:800]}", flush=True)

        event        = data.get("event", {})
        item_id      = event.get("pulseId") or event.get("itemId")
        body         = event.get("textBody", "") or event.get("body", "")
        creator_name = event.get("userName", "")
        user_id      = event.get("userId", "")

        print(f"Parsed: item_id={item_id} creator='{creator_name}' user_id={user_id} body_preview={body[:100]}", flush=True)

        if not item_id:
            print("SKIP: no item_id", flush=True)
            return {"status": "skip - no item"}, 200

        if not body:
            print("SKIP: no body", flush=True)
            return {"status": "skip - no body"}, 200

        if "AI_Assist" in creator_name or "ai_assist" in creator_name.lower():
            print(f"SKIP: self-mention from {creator_name}", flush=True)
            return {"status": "skip - self"}, 200

        # Более либеральная проверка mention — ловим в любом виде:
        # "@AI_Assist", "@ai_assist", или если есть mention объект в event
        body_lower = body.lower()
        has_mention = (
            "ai_assist" in body_lower or
            "aiassist" in body_lower or
            "ai-assist" in body_lower or
            event.get("mentionedUsers") or
            event.get("mentioned_users")
        )

        if not has_mention:
            print(f"SKIP: no mention in body. Body was: {body[:200]}", flush=True)
            return {"status": "skip - no mention"}, 200

        print(f"PROCESSING item {item_id}", flush=True)

        threading.Thread(
            target=process_monday_async,
            args=(item_id, body, str(user_id)),
            daemon=True
        ).start()

        return {"status": "ok"}, 200

    except Exception as e:
        print(f"Monday webhook error: {e}")
        return {"status": "error"}, 500


# ─────────────────────────────────────────
# MONDAY BACKGROUND PROCESSOR
# ─────────────────────────────────────────

def process_monday_async(item_id, body, user_id=""):
    try:
        from monday_webhook import post_monday_comment, get_item_details
        from agent import process_monday_request

        print(f"Processing Monday item {item_id} from user {user_id}")

        item_details = get_item_details(item_id)
        items        = item_details.get("data", {}).get("items", [])
        if not items:
            print(f"Item {item_id} not found in Monday")
            return

        item        = items[0]
        client_name = item.get("name", "Client")
        cols        = {c["id"]: c["text"] for c in item.get("column_values", [])}

        # ─── Правильные column IDs для нового Main board ───
        file_number = cols.get("text5__1", "")       # колонка #
        job_type    = cols.get("text6__1", "")       # Job description
        location    = cols.get("location__1", "")    # Location (was text__1 — это было Notes!)
        notes       = cols.get("text__1", "")        # Notes
        phone       = cols.get("phone__1", "")
        email       = cols.get("email__1", "")
        dropbox_url = cols.get("link__1", "")

        print(f"Parsed from Monday: client='{client_name}' number='{file_number}' "
              f"job='{job_type}' location='{location}' phone='{phone}' email='{email}'")

        clean_body = (body
                      .replace("@AI_Assist", "")
                      .replace("@ai_assist", "")
                      .strip())

        mention = (
            f'<a class="user_mention_editor" '
            f'data-mention-type="User" '
            f'data-mention-id="{user_id}">@mention</a> '
            if user_id else ""
        )

        # Monday всегда на отдельном loop — не блокируем Telegram
        loop = asyncio.new_event_loop()
        try:
            response = loop.run_until_complete(
                process_monday_request(
                    item_id=item_id,
                    client_name=client_name,
                    file_number=file_number,
                    job_type=job_type,
                    location=location,
                    notes=notes,
                    phone=phone,
                    email=email,
                    dropbox_url=dropbox_url,
                    message=clean_body
                )
            )
        finally:
            loop.close()

        result = post_monday_comment(item_id, f"{mention}{response}")
        print(f"Monday comment posted for item {item_id}: {result}")

    except Exception as e:
        print(f"process_monday_async error for item {item_id}: {e}")
        traceback.print_exc()


# ─────────────────────────────────────────
# BOT STARTUP
# ─────────────────────────────────────────

async def run_bot():
    global bot_app
    bot_app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    bot_app.add_handler(CommandHandler("start", start))
    bot_app.add_handler(CommandHandler("new", new_estimate))
    bot_app.add_handler(CommandHandler("dropbox", handle_dropbox))
    bot_app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    bot_app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    await bot_app.initialize()
    await bot_app.bot.set_webhook(WEBHOOK_URL)
    await bot_app.start()
    print("Webhook установлен:", WEBHOOK_URL)
    bot_ready.set()


def start_bot_thread():
    global bot_loop
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    bot_loop = loop
    loop.run_until_complete(run_bot())
    loop.run_forever()


if __name__ == "__main__":
    t = threading.Thread(target=start_bot_thread, daemon=True)
    t.start()

    bot_ready.wait(timeout=30)
    if not bot_ready.is_set():
        print("WARNING: Bot did not initialize in time")

    port = int(os.getenv("PORT", 8080))
    print(f"Flask запущен на порту {port}")
    app.run(host="0.0.0.0", port=port, use_reloader=False)
