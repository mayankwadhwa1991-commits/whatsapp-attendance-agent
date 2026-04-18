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
import openpyxl

def write_to_excel(extracted):
    wb = openpyxl.load_workbook("template.xlsx")
    ws = wb["Attendance"]

    # 1. Extract date number (e.g., 18 from 18/04/2026)
    date_line = extracted["date"]
    date_num = int(date_line.split(":")[1].strip().split("/")[0])

    # 2. Find the correct date column in row 4
    date_col = None
    for col in range(1, ws.max_column + 1):
        cell_value = ws.cell(row=4, column=col).value
        if cell_value == date_num:
            date_col = col
            break

    if not date_col:
        print("DATE COLUMN NOT FOUND IN TEMPLATE")
        return

    # P/A column = date_col
    # OT column = date_col + 1
    pa_col = date_col
    ot_col = date_col + 1

    # 3. Write attendance for each employee
    for emp in extracted["employees"]:
        # Example: "NR025 ARUN = P+2"
        code = emp.split()[0]  # NR025
        value = emp.split("=")[1].strip()  # P+2 or A or P--

        # Split attendance and OT
        if "+" in value:
            att, ot = value.split("+")
        elif "-" in value:
            att, ot = value.split("-")
        else:
            att, ot = value, "0"

        att = att.strip()
        ot = ot.strip()

        # Find employee row
        emp_row = None
        for row in range(1, ws.max_row + 1):
            if ws.cell(row=row, column=1).value == code:
                emp_row = row
                break

        if not emp_row:
            print(f"Employee {code} not found in template")
            continue

        # Write values
        ws.cell(row=emp_row, column=pa_col).value = att
        ws.cell(row=emp_row, column=ot_col).value = ot

    # Save output file
    wb.save("attendance_output.xlsx")
    print("Excel updated successfully!")
