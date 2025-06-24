import os
import pandas as pd
from fpdf import FPDF
from fpdf.enums import XPos, YPos
import yagmail
from datetime import datetime, timedelta
import requests

# ─── CONFIG ─────────────────────────────────────────────────
GOOGLE_SHEET_CSV_URL = (
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vTxaUq2leO_eZIQWMWzeSEtBbj0tknrnkhLInZjND3MfkRgZ77qBgWPVnDm6w-rEUFbt5pp5dBTyMLD/pub?output=csv"
)
WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbwsdHMmQ3t5Qa89GhmZ4CXqeRRDxmNnH-lY_3GcrBnhu5hGrkv88qdptjx0HBil2bCKpQ/exec"

OUTPUT_DIR = r"C:\Users\academytraining\Downloads\HARI-CHARLEZ\generated_pds"
FONT_PATH = r"C:\Users\academytraining\Downloads\HARI-CHARLEZ\DejaVuSans.ttf"
FONT_NAME = "DejaVu"
EMAIL_SENDER = "yogeshdark2527@gmail.com"
APP_PASSWORD = "gpzb hszg thff mevr"
LOGO_PATH = r"C:\Users\academytraining\Downloads\HARI-CHARLEZ\logo.png"
CORNER_IMG_PATH = r"C:\Users\academytraining\Downloads\HARI-CHARLEZ\corner.png"
CORNER_IMG_PATH_1 = r"C:\Users\academytraining\Downloads\HARI-CHARLEZ\corner1.png"
SIGN_IMG = r"C:\Users\academytraining\Downloads\HARI-CHARLEZ\sign.png"
SENT_EMAILS_LOG = os.path.join(OUTPUT_DIR, "sent_emails.txt")

# ─── PREPARE ENV ────────────────────────────────────────────
os.makedirs(OUTPUT_DIR, exist_ok=True)
df = pd.read_csv(GOOGLE_SHEET_CSV_URL)
df.columns = df.columns.str.strip()

# Load already sent email addresses
if os.path.exists(SENT_EMAILS_LOG):
    with open(SENT_EMAILS_LOG, 'r') as f:
        sent_emails = set(line.strip() for line in f if line.strip())
else:
    sent_emails = set()

newly_sent = []
already_sent = []

def write_row(pdf, label1, value1, label2=None, value2=None,
              bold_font=FONT_NAME, regular_font=FONT_NAME, fs=9, lh=6,
              w1=42, w2=53, w3=42, w4=53):
    page_w = pdf.w - pdf.l_margin - pdf.r_margin
    gray = 230
    if label2 is None:
        pdf.set_draw_color(gray, gray, gray)
        pdf.set_font(bold_font, 'B', fs)
        pdf.cell(w1, lh, label1, border=1)
        pdf.set_font(regular_font, '', fs)
        pdf.cell(page_w - w1, lh, value1, border=1)
        pdf.ln(lh)
    else:
        pdf.set_draw_color(gray, gray, gray)
        pdf.set_font(bold_font, 'B', fs)
        pdf.cell(w1, lh, label1, border=1)
        pdf.set_font(regular_font, '', fs)
        pdf.cell(w2, lh, value1, border=1)
        pdf.set_font(bold_font, 'B', fs)
        pdf.cell(w3, lh, label2, border=1)
        pdf.set_font(regular_font, '', fs)
        pdf.cell(w4, lh, value2, border=1)
        pdf.ln(lh)

# ─── Main Loop ───────────────────────────
for _, row in df.iterrows():
    user = {
        "Name":       row.get("Name", ""),
        "RegNo":      row.get("Register Number", ""),
        "Deg":        row.get("Degree",""),
        "EnrollID":   row.get("Enrollment ID", ""),
        "Technology": row.get("Technology", ""),
        "Domain":     row.get("Domain", ""),
        "College":    row.get("College Name", ""),
        "Course":     row.get("Course", ""),
        "StartDate":  row.get("Start Date", ""),
        "EndDate":    row.get("End Date", ""),
        "Phone":      row.get("PhoneNumber", ""),
        "Stipend":    row.get("Stipend", "") or "Not Applicable",
        "Email":      row.get("Email", ""),
    }

    pdf = FPDF(format="A4")
    pdf.add_page()
    pdf.set_left_margin(15)
    pdf.set_right_margin(15)
    pdf.set_auto_page_break(auto=True, margin=20)

    if os.path.exists(FONT_PATH):
        pdf.add_font(FONT_NAME, '', FONT_PATH, uni=True)
        pdf.add_font(FONT_NAME, 'B', FONT_PATH, uni=True)
        pdf.set_font(FONT_NAME, '', 9)
    else:
        pdf.set_font("Arial", size=9)

    if os.path.exists(LOGO_PATH):
        pdf.image(LOGO_PATH, x=10, y=5, w=60)

    if os.path.exists(SIGN_IMG):
        pdf.image(SIGN_IMG, x=10, y=230, w=50)

    send_date = datetime.now().strftime("%d-%b-%Y")
    next_date = (datetime.now() + timedelta(days=1)).strftime("%d-%b-%Y")

    pdf.set_xy(pdf.w - 50, 30)
    pdf.set_font(pdf.font_family, '', 8)
    pdf.cell(0, 10, f"Date: {send_date}", ln=True)

    pdf.set_font(pdf.font_family, '', 9)
    pdf.set_xy(15, 50)
    pdf.cell(0, 6, f"Dear {user['Name']},", ln=True)
    pdf.cell(0, 6, f"{user['RegNo']}", ln=True)
    pdf.set_font(pdf.font_family, 'B', 9)
    pdf.cell(0, 6, f"{user['Deg']}", ln=True)
    pdf.cell(0, 6, user["College"], ln=True)
    pdf.ln(2)

    pdf.set_font(pdf.font_family, '', 9)
    pdf.multi_cell(0, 6,
        "Congratulations!\n"
        "We are pleased to offer you an opportunity to undergo On-The-Job Training (OJT) at VDart Academy. "
        "This training program is designed to provide you with practical exposure and hands-on experience, "
        "enhancing your skills and preparing you for future career opportunities.\n\n"
    )

    pdf.set_font(pdf.font_family, 'B', 10)
    pdf.cell(0, 6, "Internship Details:", ln=True)
    pdf.set_font(pdf.font_family, '', 9)
    pdf.ln(1)

    write_row(pdf, "Enrollment:", "Academic Internship", "Enrollment ID:", user["EnrollID"])
    write_row(pdf, "Technology:", user["Technology"], "Domain:", user["Domain"])
    write_row(pdf, "Organization:", "VDart Academy", "Location:", "GCE – Mannarpuram")
    write_row(pdf, "Start Date:", user["StartDate"], "End Date:", user["EndDate"])
    write_row(pdf, "Stipend:", user["Stipend"], "Reporting SME:", "Anubharathi P")
    write_row(pdf, "Shift Time:", "9:00 AM to 1:00 PM IST", "Shift Days:", "Monday – Friday")
    write_row(pdf, "SME Email:", "anubharathi.p@vdartinc.com", "SME Mobile:", "+91 99445 48333")
    pdf.ln(4)

    pdf.multi_cell(0, 6,
        f"\nWe believe this opportunity will contribute to your professional development, and we look forward to your active participation. "
        f"Kindly confirm your acceptance by signing a copy of this letter by {next_date}.\n\n"
        "Should you have any questions, feel free to contact us.\n"
    )
    pdf.ln(2)
    pdf.cell(0, 6, user["Course"], ln=True)

    # Footer
    pdf.set_font(pdf.font_family, 'B', 10)
    pdf.set_xy(145, 240)
    pdf.cell(0, 6, "+91 99445 48333", ln=True)
    pdf.set_xy(145, 245)
    pdf.cell(0, 6, "info@vdartacademy.com", ln=True)
    pdf.set_xy(145, 250)
    pdf.cell(0, 6, "Mannarpuram,Trichy-620020", ln=True)

    if os.path.exists(CORNER_IMG_PATH):
        pdf.image(CORNER_IMG_PATH, x=pdf.w - 50, y=0, w=50, h=20)
    if os.path.exists(CORNER_IMG_PATH_1):
        pdf.image(CORNER_IMG_PATH_1, x=0, y=pdf.h - 20, w=50, h=20)

    fname = f"OfferLetter_{user['Name'].replace(' ', '_').replace('/', '_')}.pdf"
    fpath = os.path.join(OUTPUT_DIR, fname)
    pdf.output(fpath)
    print("✅ Created:", fname)

    if user["Email"]:
        if user["Email"] not in sent_emails:
            try:
                yag = yagmail.SMTP(EMAIL_SENDER, APP_PASSWORD)
                yag.send(
                    to=user["Email"],
                    subject="VDart Academy Internship Offer Letter",
                    contents=(
                        f"Dear {user['Name']},\n\n"
                        "Please find your offer letter attached.\n\n"
                        "Best regards,\nVDart Academy"
                    ),
                    attachments=fpath
                )
                print("📩 Sent to:", user["Email"])
                newly_sent.append(user["Name"])
                with open(SENT_EMAILS_LOG, 'a') as f:
                    f.write(user["Email"] + "\n")
                sent_emails.add(user["Email"])

                # 📢 Update Mail Status in Google Sheet
                requests.post(WEBHOOK_URL, data={"email": user["Email"], "status": "Mail Sent"})

            except Exception as e:
                print("❌ Email failed:", e)
                requests.post(WEBHOOK_URL, data={"email": user["Email"], "status": "Failed"})
        else:
            print("⏭️ Already sent to:", user["Email"])
            already_sent.append(user["Name"])
    else:
        print("⚠️ No email for:", user["Name"])

# ─── Summary ───────────────────────────
print("\n✅ Newly Sent Offer Letters:")
for name in newly_sent:
    print("  ➤", name)

print("\n⏭️ Already Sent Previously:")
for name in already_sent:
    print("  ➤", name)

print("\n🎉 All done!")
