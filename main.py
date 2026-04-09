import os
import httpx
import json
from dotenv import load_dotenv
from flask import Flask, request, Response
from twilio.twiml.messaging_response import MessagingResponse
from telegram import Update, Bot
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
import asyncio

load_dotenv()

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
WEBHOOK_URL = "https://estimator-agent-production.up.railway.app/telegram"

app = Flask(__name__)
bot_app = None

@app.route("/sms", methods=["POST"])
def sms_reply():
    from agent import process_request
    phone = request.form.get("From", "")
    body = request.form.get("Body", "").strip()
    print(f"SMS from {phone}: {body}")
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
    global bot_app
    if bot_app:
        data = request.get_json(force=True)
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        update = Update.de_json(data, bot_app.bot)
        loop.run_until_complete(bot_app.process_update(update))
        loop.close()
    return Response(status=200)

@app.route("/health", methods=["GET"])
def health():
    return "OK"

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Привет! Я Estimator Agent.\n\nОпиши работу или скинь фото.\n\n/new — новый эстимейт")

async def new_estimate(update: Update, context: ContextTypes.DEFAULT_TYPE):
    from agent import clear_conversation
    user_id = update.message.from_user.id
    result = clear_conversation(user_id)
    await update.message.reply_text(result)

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

async def setup_bot():
    global bot_app
    bot_app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    bot_app.add_handler(CommandHandler("start", start))
    bot_app.add_handler(CommandHandler("new", new_estimate))
    bot_app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    bot_app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    await bot_app.initialize()
    await bot_app.bot.set_webhook(WEBHOOK_URL)
    await bot_app.start()
    print("Webhook установлен:", WEBHOOK_URL)

if __name__ == "__main__":
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    loop.run_until_complete(setup_bot())
    port = int(os.getenv("PORT", 8080))
    print(f"Flask запущен на порту {port}")
    app.run(host="0.0.0.0", port=port, use_reloader=False)
