import os
import requests
from dotenv import load_dotenv

load_dotenv()

MONDAY_AI_API_KEY = os.getenv("MONDAY_AI_API_KEY")
MONDAY_API_KEY = os.getenv("MONDAY_API_KEY")


def post_monday_comment(item_id, message):
    query = """mutation ($item: ID!, $body: String!) {
        create_update(item_id: $item, body: $body) { id }
    }"""
    headers = {
        "Authorization": MONDAY_AI_API_KEY,
        "Content-Type": "application/json",
        "API-Version": "2023-10"
    }
    r = requests.post("https://api.monday.com/v2", headers=headers, json={
        "query": query,
        "variables": {"item": str(item_id), "body": message}
    })
    print(f"post_monday_comment {item_id}: {r.status_code} {r.text[:200]}")
    return r.json()


def get_item_details(item_id):
    query = """query ($item: [ID!]) {
        items(ids: $item) {
            id
            name
            column_values {
                id
                text
            }
            updates {
                id
                body
                created_at
                creator { name }
            }
        }
    }"""
    headers = {
        "Authorization": MONDAY_API_KEY,
        "Content-Type": "application/json",
        "API-Version": "2023-10"
    }
    r = requests.post("https://api.monday.com/v2", headers=headers, json={
        "query": query,
        "variables": {"item": [str(item_id)]}
    })
    print(f"get_item_details {item_id}: {r.status_code} {r.text[:200]}")
    return r.json()
