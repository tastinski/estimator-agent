import os
import httpx
import json
from dotenv import load_dotenv
from flask import Flask, request, Response
from twilio.twiml.messaging_response import MessagingResponse
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
import asyncio
import threading

load_dotenv()

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
WEBHOOK_URL = "https://estimator-agent-production.up.railway.app/telegram"

app = Flask(__name__)
bot_app = None
bot_loop = None

@app.route("/sms", methods=["POST"])
def sms_reply():
    from agent import process_request
    phone = request.form.get("From", "")
    body = request.form.get("Body", "").strip()
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    result = loop.run_until_complete(process_request(phone, body))
    loop.close()
    resp = MessagingResponse()
    if len(result) > 1600:
        result = result[:1597] + "..."
    resp.message(result)
    return str(resp)

@app.route("/telegram", methods=["POST"])
def telegram_webhook():
    global bot_app, bot_loop
    if bot_app and bot_loop:
        data = request.get_json(force=True)
        update = Update.de_json(data, bot_app.bot)
        asyncio.run_coroutine_threadsafe(bot_app.process_update(update), bot_loop)
    return Response(status=200)

@app.route("/health", methods=["GET"])
def health():
    return "OK"

@app.route("/monday_webhook", methods=["POST"])
def monday_webhook():
    try:
        data = request.get_json(force=True)
        if "challenge" in data:
            return {"challenge": data["challenge"]}, 200
        event = data.get("event", {})
        item_id = event.get("pulseId") or event.get("itemId")
        body = event.get("textBody", "") or event.get("body", "")
        creator_name = event.get("userName", "")
        if not item_id or not body:
            return {"status": "skip"}, 200
        if "AI_Assist" in creator_name or "ai_assist" in creator_name.lower():
            return {"status": "skip"}, 200
        if "@AI_Assist" not in body and "@ai_assist" not in body.lower():
            return {"status": "skip"}, 200
        from monday_webhook import post_monday_comment, get_item_details
        item_details = get_item_details(item_id)
        items = item_details.get("data", {}).get("items", [])
        if not items:
            return {"status": "item not found"}, 200
        item = items[0]
        client_name = item.get("name", "Client")
        cols = {c["id"]: c["text"] for c in item.get("column_values", [])}
        job_type = cols.get("text6__1", "")
        location = cols.get("text__1", "")
        context = f"Project: {client_name}. Job: {job_type}. Location: {location}. Question: {body}"
        from agent import process_request
        global bot_loop
        if bot_loop:
            future = asyncio.run_coroutine_threadsafe(
                process_request(f"monday_{item_id}", context), bot_loop
            )
            response = future.result(timeout=60)
        else:
            loop = asyncio.new_event_loop()
            response = loop.run_until_complete(process_request(f"monday_{item_id}", context))
            loop.close()
        post_monday_comment(item_id, f"**AI_Assist:** {response}")
        return {"status": "ok"}, 200
    except Exception as e:
        print(f"Monday webhook error: {e}")
        return {"status": "error"}, 500

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Привет! Я Estimator Agent.\n\nОпиши работу или скинь фото.\n\n/new — новый эстимейт\n/dropbox [имя клиента] — анализ фото из Dropbox")

async def new_estimate(update: Update, context: ContextTypes.DEFAULT_TYPE):
    from agent import clear_conversation
    user_id = update.message.from_user.id
    result = clear_conversation(user_id)
    await update.message.reply_text(result)

async def handle_dropbox(update: Update, context: ContextTypes.DEFAULT_TYPE):
    from agent import process_dropbox_photos
    user_id = update.message.from_user.id
    client_name = " ".join(context.args) if context.args else ""
    if not client_name:
        await update.message.reply_text("Напиши имя клиента: /dropbox Maria Gonzalez")
        return
    await update.message.reply_text(f"📸 Ищу фото в Dropbox для {client_name}... Это займёт 1-2 минуты.")
    try:
        result = await process_dropbox_photos(user_id, client_name)
        chunk_size = 4000
        for i in range(0, len(result), chunk_size):
            await update.message.reply_text(result[i:i+chunk_size])
    except Exception as e:
        await update.message.reply_text(f"❌ Ошибка: {str(e)}")

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    from agent import process_request
    user_id = update.message.from_user.id
    user_message = update.message.text
    await update.message.reply_text("Обрабатываю...")
    result = await process_request(user_id, user_message)
    await update.message.reply_text(result)

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    from agent import process_photo
    user_id = update.message.from_user.id
    caption = update.message.caption or ""
    await update.message.reply_text("Фото получено, анализирую...")
    photo = update.message.photo[-1]
    file = await context.bot.get_file(photo.file_id)
    async with httpx.AsyncClient() as client:
        response = await client.get(file.file_path)
        image_data = response.content
    result = await process_photo(user_id, image_data, caption)
    await update.message.reply_text(result)

async def run_bot():
    global bot_app, bot_loop
    bot_loop = asyncio.get_event_loop()
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

def start_bot_thread():
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    global bot_loop
    bot_loop = loop
    loop.run_until_complete(run_bot())
    loop.run_forever()

if __name__ == "__main__":
    t = threading.Thread(target=start_bot_thread, daemon=True)
    t.start()
    import time
    time.sleep(3)
    port = int(os.getenv("PORT", 8080))
    print(f"Flask запущен на порту {port}")
    app.run(host="0.0.0.0", port=port, use_reloader=False)
