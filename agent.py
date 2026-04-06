import os
from anthropic import Anthropic
from dotenv import load_dotenv
from rag import search_prices, init_database

load_dotenv()

client = Anthropic(api_key=os.getenv("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"))

# История разговоров по каждому пользователю
conversations = {}

init_database()

async def process_request(user_id: int, user_message: str) -> str:
    try:
        # Инициализируем историю если новый пользователь
        if user_id not in conversations:
            conversations[user_id] = []

        # Добавляем сообщение пользователя в историю
        conversations[user_id].append({
            "role": "user",
            "content": user_message
        })

        # Ищем цены в RAG
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

When ready to estimate, use this format:

CLIENT: [name]
JOB: [job type]
LOCATION: [location]

BREAKDOWN:
- [item 1]: $[price]
- [item 2]: $[price]
- [item 3]: $[price]

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

        # Сохраняем ответ агента в историю
        conversations[user_id].append({
            "role": "assistant",
            "content": assistant_message
        })

        return assistant_message

    except Exception as e:
        return f"❌ Ошибка: {str(e)}"

def clear_conversation(user_id: int):
    if user_id in conversations:
        del conversations[user_id]
    return "🔄 Разговор сброшен. Начнём заново!"
