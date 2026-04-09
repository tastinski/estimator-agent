import os
import sys
import fcntl
import httpx
import threading
from dotenv import load_dotenv
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
from twilio_handler import app as flask_app

load_dotenv()

lock_file = open("/tmp/estimator_bot.lock", "w")
try:
    fcntl.flock(lock_file, fcntl.LOCK_EX | fcntl.LOCK_NB)
except IOError:
    print("Другой инстанс уже запущен. Выходим.")
    sys.exit(0)

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")

def run_flask():
    port = int(os.getenv("PORT", 5000))
    flask_app.run(host="0.0.0.0", port=port, use_reloader=False)

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

if __name__ == "__main__":
    flask_thread = threading.Thread(target=run_flask, daemon=True)
    flask_thread.start()
    print("Flask запущен на порту", os.getenv("PORT", 5000))
    
    app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("new", new_estimate))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    print("Бот запущен!")
    app.run_polling(drop_pending_updates=True, allowed_updates=Update.ALL_TYPES)
