import gspread
from oauth2client.service_account import ServiceAccountCredentials
from fpdf import FPDF
from datetime import datetime, timedelta
import yagmail
import os

# ─── CONFIG ───
CORNER_IMG_PATH = r"C:\Users\Yogesh\Downloads\HARI_CHARLEZ\final-step\corner.png"
CORNER_IMG_PATH_1 = r"C:\Users\Yogesh\Downloads\HARI_CHARLEZ\final-step\corner1.png"
OUTPUT_DIR = r"C:\Users\Yogesh\Downloads\HARI_CHARLEZ\final-step\formatted_pdfs"
FONT_PATH = r"C:\Users\Yogesh\Downloads\HARI_CHARLEZ\final-step\DejaVuSans.ttf"
LOGO_PATH = r"C:\Users\Yogesh\Downloads\HARI_CHARLEZ\final-step\logo.png"
SIGN_IMG = r"C:\Users\Yogesh\Downloads\HARI_CHARLEZ\final-step\sign.png"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ─── Google Sheet Auth ───
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)
sheet = client.open("Final Sheet").sheet1

expected_headers = [
    'No', 'Emp ID', 'Asset Code', 'Name', 'Reg No', 'Email', 'Aadhar', 'Phone', 'School / College',
    'College Location', 'Major', 'Year', 'Department', 'Role', 'D.O.B', 'Gender', 'FG', 'Method', 'Refered From',
    'Person Status', 'Shift Timing', 'Working Days', 'Shift', 'Join Date',
    'End Date', 'Model', 'Offer Letter', 'Offer Signed', 'Completion Letter'
]

data = sheet.get_all_records(expected_headers=expected_headers)

enrollment_dict = {row['Reg No']: row for row in data}

def write_row(pdf, label1, value1, label2=None, value2=None, fs=10, lh=8, w1=40, w2=60, w3=40, w4=50):
    gray = 230
    pdf.set_draw_color(gray, gray, gray)
    pdf.set_font("Arial", 'B', fs)
    pdf.cell(w1, lh, label1, border=1)
    pdf.set_font("Arial", '', fs)
    pdf.cell(w2, lh, value1, border=1)
    if label2 and value2:
        pdf.set_font("Arial", 'B', fs)
        pdf.cell(w3, lh, label2, border=1)
        pdf.set_font("Arial", '', fs)
        pdf.cell(w4, lh, value2, border=1)
    pdf.ln(lh)

internship_base_number = 100
i = 2
for enrollment_id, student in enrollment_dict.items():
    offer_status = str(student['Offer Letter']).strip().lower()
    if offer_status == "sended":
        i += 1
        continue

    internship_id = f"A0{internship_base_number}"
    send_date = datetime.now().strftime("%d-%b-%Y")
    next_date = (datetime.now() + timedelta(days=1)).strftime("%d-%b-%Y")

    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.set_left_margin(15)
    pdf.set_right_margin(15)
    pdf.set_font("Arial", size=10)

    pdf.set_xy(pdf.w - 50, 30)
    pdf.set_font("Arial", '', 8)
    pdf.cell(0, 10, f"Date: {send_date}", ln=True)

    if os.path.exists(LOGO_PATH):
        pdf.image(LOGO_PATH, x=10, y=5, w=60)

    if os.path.exists(SIGN_IMG):
        pdf.image(SIGN_IMG, x=10, y=230, w=50)

    pdf.set_xy(15, 50)
    pdf.set_font("Arial", '', 10)
    pdf.cell(0, 6, f"Dear {student['Name']},", ln=True)
    pdf.cell(0, 6, f"{student['Reg No']}", ln=True)
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(0, 6, f"{student['Major']}", ln=True)
    pdf.cell(0, 6, student["School / College"], ln=True)
    pdf.ln(4)

    pdf.set_font("Arial", '', 10)
    pdf.multi_cell(0, 6, "Congratulations!\nWe are pleased to offer you an opportunity to undergo On-The-Job Training (OJT) at VDart Academy. This training program is designed to provide you with practical exposure and hands-on experience, enhancing your skills and preparing you for future career opportunities.\n")

    pdf.set_font("Arial", 'B', 11)
    pdf.cell(0, 8, "Internship Details:", ln=True)
    pdf.set_font("Arial", '', 10)
    pdf.ln(1)

    write_row(pdf, "Enrollment", "Academic Internship", "Enrollment ID", internship_id)
    write_row(pdf, "Technology", student["Role"], "Domain", "example")
    write_row(pdf, "Organization", "VDart Academy", "Location", "GCE - Mannapuram")
    write_row(pdf, "Start Date", student["Join Date"], "End Date", student["End Date"])
    write_row(pdf, "Stipend", "example", "Reporting SME", "Anubharathi P")
    write_row(pdf, "Shift Time", "2:00 PM to 6:00 PM IST", "Shift Days", "Monday to Friday")
    write_row(pdf, "SME Email", "anubharathi.p@vdartinc.com", "SME Mobile", "+91 99445 48333")

    pdf.ln(6)
    pdf.multi_cell(0, 6, f"We believe this opportunity will contribute to your professional development, and we look forward to your active participation. Kindly confirm your acceptance by signing a copy of this letter by {next_date}.\n\nShould you have any questions, feel free to contact us.\n")

    pdf.set_xy(145, 240)
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(0, 6, "+91 99445 48333", ln=True)
    pdf.set_xy(145, 245)
    pdf.cell(0, 6, "info@vdartacademy.com", ln=True)
    pdf.set_xy(145, 250)
    pdf.cell(0, 6, "Mannapuram,Trichy-620020", ln=True)

    if os.path.exists(CORNER_IMG_PATH):
        pdf.image(CORNER_IMG_PATH, x=pdf.w - 50, y=0, w=50, h=20)
    if os.path.exists(CORNER_IMG_PATH_1):
        pdf.image(CORNER_IMG_PATH_1, x=0, y=pdf.h - 20, w=50, h=20)

    filename = f"OfferLetter_{student['Name'].replace(' ', '_').replace('/', '_')}.pdf"
    filepath = os.path.join(OUTPUT_DIR, filename)
    pdf.output(filepath)
    print("✅ Created:", filename)

    join_date_str = student.get('Join Date')
    try:
        join_date = datetime.strptime(join_date_str.strip(), "%d/%m/%Y")
    except Exception:
        join_date = None

    scheduled_time = join_date - timedelta(days=2) if join_date else datetime.now()

    if datetime.now() >= scheduled_time:
        try:
            yag = yagmail.SMTP("vdartacademy111@gmail.com", "wnip bjez wxjb mzig")  # Your App Password
            subject = f"VDart Academy | On-the-Job Training - Onboarding Letter | {student['Name']}"
            body = f"""
                <div style="font-family: Arial, sans-serif; font-size: 14px; color: #fff;">
                    <p style="margin: 4px 0; line-height: 1.3;">Dear <b>{student['Name']}</b>
                    
                        Congratulations on being selected for the <b>On-the-Job Training (OJT)</b> in <b>{student['Role']}</b> with <b>VDart Academy</b>.
                        
                        We are pleased to inform you that your OJT will commence on <b>{join_date_str}</b>.
                        
                        Please find your onboarding letter attached. Kindly review, countersign, and share it with us to confirm your acceptance.
                        
                        We look forward to a meaningful and successful journey ahead. If you have any questions, feel free to reach out.
                        Welcome aboard!</b>
                    </p>
                </div>
            """
            yag.send(to=student['Email'], subject=subject, contents=[body, filepath])
            print(f"📧 Email sent to {student['Email']}")
            sheet.update_cell(i, 27, "Sended")
        except Exception as e:
            print(f"❌ Email failed to {student['Email']}: {e}")
            sheet.update_cell(i, 27, "Not Sended")


    i += 1
    internship_base_number += 1

print("\n🎉 All PDFs processed and emails handled!")
