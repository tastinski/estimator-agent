import os
from flask import Flask, request, Response
from twilio.twiml.messaging_response import MessagingResponse
from dotenv import load_dotenv
import asyncio

load_dotenv()

app = Flask(__name__)

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
    return Response(status=200)

@app.route("/health", methods=["GET"])
def health():
    return "OK"

@app.route("/monday_webhook", methods=["POST"])
def monday_webhook():
    import asyncio
    from agent import process_request
    from monday_webhook import post_monday_comment, get_item_details
    
    try:
        data = request.get_json(force=True)
        print(f"Monday webhook: {data}")
        
        event = data.get("event", {})
        item_id = event.get("pulseId") or event.get("itemId")
        body = event.get("textBody", "") or event.get("body", "")
        creator_name = event.get("userName", "")
        
        if not item_id or not body:
            return {"status": "skip"}, 200
        
        if "AI_Assist" in creator_name or "ai_assist" in creator_name.lower():
            return {"status": "skip - own message"}, 200
        
        if "@AI_Assist" not in body and "@ai_assist" not in body.lower():
            return {"status": "skip - not mentioned"}, 200
        
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
        
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        response = loop.run_until_complete(process_request(f"monday_{item_id}", context))
        loop.close()
        
        post_monday_comment(item_id, f"**AI_Assist:** {response}")
        
        return {"status": "ok"}, 200
    except Exception as e:
        print(f"Monday webhook error: {e}")
        import traceback
        traceback.print_exc()
        return {"status": "error", "message": str(e)}, 500
