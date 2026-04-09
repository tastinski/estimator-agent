import os
from flask import Flask, request
from twilio.twiml.messaging_response import MessagingResponse
from dotenv import load_dotenv
import asyncio

load_dotenv()

app = Flask(__name__)

conversations = {}

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
    # SMS лимит 1600 символов
    if len(result) > 1600:
        result = result[:1597] + "..."
    resp.message(result)
    return str(resp)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", 5000)))
