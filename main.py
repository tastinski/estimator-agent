import os
from dotenv import load_dotenv
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters, ContextTypes
from agent import process_request, clear_conversation

load_dotenv()

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 Привет! Я Estimator Agent.\n\n"
        "Опиши работу и я задам уточняющие вопросы,\n"
        "потом сделаю точный эстимейт.\n\n"
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

if __name__ == "__main__":
    app = ApplicationBuilder().token(TELEGRAM_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("new", new_estimate))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    print("🤖 Бот запущен!")
    app.run_polling()
