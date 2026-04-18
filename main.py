from fastapi import FastAPI, Request
import openpyxl
import os

app = FastAPI()

# -----------------------------
# SAFE MESSAGE EXTRACTION
# -----------------------------
def get_message_text(data):
    try:
        return data["entry"][0]["changes"][0]["value"]["messages"][0]["text"]["body"]
    except:
        return ""


# -----------------------------
# SAFE ATTENDANCE EXTRACTION
# -----------------------------
def extract_attendance(text):
    lines = [l.strip() for l in text.split("\n") if l.strip()]

    # Find DATE line safely
    date_candidates = [l for l in lines if "DATE" in l.upper() or "/" in l]
    if not date_candidates:
        return {"error": "No date found", "raw_text": text}

    date_line = date_candidates[0]

    # Company name = line after date
    try:
        company_line = lines[lines.index(date_line) + 1]
    except:
        company_line = "UNKNOWN COMPANY"

    # Extract employee rows
    employees = []
    for l in lines:
        if "=" in l:
            employees.append(l)

    return {
        "date": date_line,
        "company": company_line,
        "employees": employees
    }


# -----------------------------
# WEBHOOK ENDPOINT
# -----------------------------
@app.post("/webhook")
async def receive_whatsapp(request: Request):
    data = await request.json()

    # Extract message text safely
    text = get_message_text(data)

    print("RAW MESSAGE RECEIVED:")
    print(text)

    extracted = extract_attendance(text)

    print("EXTRACTED DATA:")
    print(extracted)

    # Always return 200 OK so WhatsApp does not retry
    return {"status": "received", "extracted": extracted}


# -----------------------------
# WEBHOOK VERIFICATION (GET)
# -----------------------------
@app.get("/webhook")
async def verify(request: Request):
    params = request.query_params
    mode = params.get("hub.mode")
    challenge = params.get("hub.challenge")
    token = params.get("hub.verify_token")

    VERIFY_TOKEN = os.getenv("VERIFY_TOKEN")

    if mode == "subscribe" and token == VERIFY_TOKEN:
        return int(challenge)

    return {"error": "Invalid verification token"}
