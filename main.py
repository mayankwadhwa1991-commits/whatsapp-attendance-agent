from fastapi import FastAPI, Request
import requests
import os
from openpyxl import load_workbook, Workbook
from datetime import datetime

app = FastAPI()

VERIFY_TOKEN = "my_verify_token"
WHATSAPP_TOKEN = "YOUR_WHATSAPP_TOKEN"
ADMIN_NUMBER = "+91XXXXXXXXXX"   # Replace later

@app.get("/webhook")
async def verify(request: Request):
    params = dict(request.query_params)
    if params.get("hub.mode") == "subscribe" and params.get("hub.verify_token") == VERIFY_TOKEN:
        return int(params.get("hub.challenge"))
    return "Verification failed"

@app.post("/webhook")
async def receive_whatsapp(request: Request):
    data = await request.json()

    try:
        msg = data["entry"][0]["changes"][0]["value"]["messages"][0]
        text = msg["text"]["body"]
        sender = msg["from"]
    except:
        return {"status": "ignored"}

    extracted = extract_attendance(text)

    for emp in extracted["employees"]:
        if not employee_exists(extracted["month_file"], emp["code"]):
            ask_admin_approval(emp, extracted["location"])
            return {"status": "waiting_for_approval"}

    fill_attendance(extracted)

    return {"status": "attendance_filled"}


def extract_attendance(message):
    lines = message.split("\n")

    date_line = [l for l in lines if "DATE" in l.upper() or "/" in l][0]
    date = extract_date(date_line)

    location = lines[1].strip()

    month_name = date.strftime("%B")
    year = date.year
    month_file = f"{month_name}-{year}.xlsx"

    employees = []
    for line in lines:
        if "NR" in line.upper():
            parts = line.split("=")
            if len(parts) != 2:
                continue
            left, att = parts
            att = att.strip()
            left = left.replace("-", " ")
            tokens = left.split()
            code = tokens[0].strip()
            name = " ".join(tokens[1:]).strip()

            if "+" in att:
                attendance = "P"
                overtime = int(att.split("+")[1])
            elif "--" in att:
                attendance = "P"
                overtime = 0
            else:
                attendance = att
                overtime = 0

            employees.append({
                "code": code,
                "name": name,
                "attendance": attendance,
                "overtime": overtime
            })

    return {
        "date": date,
        "location": location,
        "month_file": month_file,
        "employees": employees
    }


def extract_date(text):
    t = text.upper().replace("DATE", "").replace(":", "").replace("-", "/")
    t = "".join(ch for ch in t if ch.isdigit() or ch == "/")
    if len(t.split("/")[-1]) == 2:
        return datetime.strptime(t.strip(), "%d/%m/%y")
    return datetime.strptime(t.strip(), "%d/%m/%Y")


def employee_exists(month_file, code):
    if not os.path.exists(month_file):
        return False
    wb = load_workbook(month_file)
    ws = wb.active
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] == code:
            return True
    return False


def ask_admin_approval(emp, location):
    msg = (
        f"New employee detected:\n"
        f"Code: {emp['code']}\n"
        f"Name: {emp['name']}\n"
        f"Location: {location}\n\n"
        f"Reply YES to add or NO to ignore."
    )
    send_whatsapp(ADMIN_NUMBER, msg)


def send_whatsapp(number, message):
    url = "https://graph.facebook.com/v17.0/YOUR_PHONE_NUMBER_ID/messages"
    headers = {
        "Authorization": f"Bearer {WHATSAPP_TOKEN}",
        "Content-Type": "application/json"
    }
    data = {
        "messaging_product": "whatsapp",
        "to": number,
        "text": {"body": message}
    }
    requests.post(url, headers=headers, json=data)


def fill_attendance(data):
    date = data["date"]
    col_base = (date.day - 1) * 2 + 4

    if not os.path.exists(data["month_file"]):
        create_month_file(data["month_file"])

    wb = load_workbook(data["month_file"])
    ws = wb.active

    for emp in data["employees"]:
        row = find_or_add_employee(ws, emp["code"], emp["name"], data["location"])
        ws.cell(row=row, column=col_base).value = emp["attendance"]
        ws.cell(row=row, column=col_base + 1).value = emp["overtime"]

    wb.save(data["month_file"])


def create_month_file(filename):
    wb = Workbook()
    ws = wb.active
    headers = ["E. No", "Employee Name", "Location"]
    for d in range(1, 32):
        headers.append(f"{d} P/A")
        headers.append(f"{d} O.T")
    ws.append(headers)
    wb.save(filename)


def find_or_add_employee(ws, code, name, location):
    for row in range(2, ws.max_row + 1):
        if ws.cell(row=row, column=1).value == code:
            return row

    new_row = ws.max_row + 1
    ws.cell(row=new_row, column=1).value = code
    ws.cell(row=new_row, column=2).value = name
    ws.cell(row=new_row, column=3).value = location
    return new_row
