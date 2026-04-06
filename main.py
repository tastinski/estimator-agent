import os
import httpx
from dotenv import load_dotenv
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
from agent import process_request, process_photo, clear_conversation

load_dotenv()

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 Привет! Я Estimator Agent.\n\n"
        "Опиши работу или скинь фото — задам уточняющие вопросы и сделаю эстимейт.\n\n"
        "Например:\n"
        "John Smith, bathroom remodel, Brooklyn NY\n\n"
        "Команды:\n"
        "/new — начать новый эстимейт"
    )

async def new_estimate(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.message.from_user.id
    result = clear_conversation(user_id)
    await update.message.reply_text(result)

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.message.from_user.id
    user_message = update.message.text
    await update.message.reply_text("⏳ Обрабатываю...")
    result = await process_request(user_id, user_message)
    await update.message.reply_text(result)

async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.message.from_user.id
    caption = update.message.caption or ""
    await update.message.reply_text("📸 Фото получено, анализирую...")

    photo = update.message.photo[-1]
    file = await context.bot.get_file(photo.file_id)
    
    async with httpx.AsyncClient() as client:
        response = await client.get(file.file_path)
        image_data = response.content

    result = await process_photo(user_id, image_data, caption)
    await update.message.reply_text(result)

if __name__ == "__main__":
    app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("new", new_estimate))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    print("🤖 Бот запущен!")
    app.run_polling()
