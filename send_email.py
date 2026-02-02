import streamlit as st
import smtplib
from email.message import EmailMessage
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from dotenv import load_dotenv
import os
import time

# --- Page Config ---
st.set_page_config(page_title="Cold Email Automation", page_icon="📧")

st.title("📧 Cold Email Automation")
st.markdown("Connect your Google Sheet and send personalized emails instantly.")

# --- 1. Load Credentials ---
load_dotenv()
LOGIN_EMAIL = os.getenv("EMAIL_USER")
LOGIN_PASSWORD = os.getenv("EMAIL_PASSWORD")

# Check for environment variables immediately
if not LOGIN_EMAIL or not LOGIN_PASSWORD:
    st.error("❌ ERROR: 'EMAIL_USER' and 'EMAIL_PASSWORD' are missing from your .env file.")
    st.stop()

# --- 2. User Inputs ---
with st.container():
    st.subheader("1️⃣ Configuration")
    sheet_url = st.text_input("🔗 Google Sheet URL", placeholder="https://docs.google.com/spreadsheets/d/...")
    
    st.subheader("2️⃣ Email Content")
    email_subject = st.text_input("📝 Subject")
    email_body = st.text_area("✍️ Body of Email", height=200, 
                              placeholder="Hi {name},\n\nI saw that {business} is doing great work...\n\nBest,\n[Your Name]")
    st.caption("💡 Tip: Use `{name}` and `{business}` as placeholders.")

# --- 3. The Sending Logic ---
if st.button("🚀 Send Emails", type="primary"):
    
    # Basic Validation
    if not sheet_url:
        st.warning("⚠️ Please enter a Google Sheet URL.")
        st.stop()
    if not email_subject or not email_body:
        st.warning("⚠️ Please enter both a Subject and a Body.")
        st.stop()

    status_area = st.empty()  # Placeholder for status updates
    log_area = st.container() # Container for the scrolling log

    # --- Connect to Google Sheets ---
    with status_area:
        st.info("🔌 Connecting to Google Sheets...")
    
    SCOPE = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.file",
        "https://www.googleapis.com/auth/drive"
    ]
    CREDS_FILE = "credentials.json"

    try:
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDS_FILE, SCOPE)
        client = gspread.authorize(creds)
        spreadsheet = client.open_by_url(sheet_url)
        sheet = spreadsheet.sheet1
    except Exception as e:
        st.error(f"❌ Google Sheets Error: {e}")
        st.stop()

    # --- Scan Columns ---
    all_data = sheet.get_all_values()
    name_idx = email_idx = status_idx = business_idx = -1
    header_row_num = 0

    # Logic to find headers
    for i, row in enumerate(all_data):
        row_lower = [str(cell).strip().lower() for cell in row]
        if "name" in row_lower and "email" in row_lower:
            name_idx = row_lower.index("name")
            email_idx = row_lower.index("email")
            status_idx = row_lower.index("status") if "status" in row_lower else len(row)
            business_idx = row_lower.index("business name") if "business name" in row_lower else -1
            header_row_num = i + 1
            break

    if name_idx == -1 or email_idx == -1:
        st.error("❌ Could not find 'Name' and 'Email' columns in the sheet.")
        st.stop()

    # --- Prepare Records ---
    records = []
    for i, row in enumerate(all_data[header_row_num:], start=header_row_num + 1):
        if len(row) > max(name_idx, email_idx) and row[email_idx].strip():
            records.append({
                "Name": row[name_idx],
                "Email": row[email_idx],
                "Business": row[business_idx] if business_idx != -1 and len(row) > business_idx else "",
                "Status": row[status_idx] if len(row) > status_idx else "",
                "row_number": i
            })

    if not records:
        st.warning("⚠️ No contacts found to process.")
        st.stop()
    
    st.success(f"✅ Found {len(records)} contacts.")

    # --- Connect to SMTP ---
    try:
        with status_area:
            st.info(f"🔐 Logging in as {LOGIN_EMAIL}...")
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(LOGIN_EMAIL, LOGIN_PASSWORD)
    except Exception as e:
        st.error(f"❌ SMTP Login Error: {e}")
        st.stop()

    # --- Send Loop ---
    progress_bar = st.progress(0)
    
    for idx, record in enumerate(records):
        name = record["Name"].strip()
        email = record["Email"].strip()
        business = record["Business"].strip()
        status = str(record["Status"]).strip().lower()
        row_num = record["row_number"]

        # Update Progress Bar
        progress_bar.progress((idx + 1) / len(records))

        if status == "sent":
            with log_area:
                st.info(f"⏭ Skipping {name} (Already sent)")
            continue

        # Send Email
        with status_area:
            st.write(f"📨 Sending to **{name}**...")

        body = email_body.replace("{name}", name).replace("{business}", business)
        msg = EmailMessage()
        msg["From"] = LOGIN_EMAIL
        msg["To"] = email
        msg["Subject"] = email_subject
        msg.set_content(body)

        try:
            server.send_message(msg)
            # Update Sheet
            sheet.update_cell(row_num, status_idx + 1, "Sent")
            with log_area:
                st.success(f"✅ Sent to {name} ({email})")
        except Exception as e:
            sheet.update_cell(row_num, status_idx + 1, "Failed")
            with log_area:
                st.error(f"❌ Failed to send to {name}: {e}")

        time.sleep(1.5) # Anti-spam delay

    server.quit()
    status_area.success("🎉 All tasks finished!")
    st.balloons()