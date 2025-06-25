import gspread
from oauth2client.service_account import ServiceAccountCredentials
from docxtpl import DocxTemplate
import yagmail
from datetime import datetime, timedelta
import os
import comtypes.client

# ─── PATHS ───────────────────────────────────────────────────────
TEMPLATE_PATH = r"C:\Users\Yogesh\Downloads\HARI_CHARLEZ\final-step\pdf-template.docx"
DOCX_OUTPUT = r"C:\Users\Yogesh\Downloads\HARI_CHARLEZ\final-step\formatted\docx"
PDF_OUTPUT = r"C:\Users\Yogesh\Downloads\HARI_CHARLEZ\final-step\formatted\pdf"

os.makedirs(DOCX_OUTPUT, exist_ok=True)
os.makedirs(PDF_OUTPUT, exist_ok=True)

# ─── PDF CONVERSION FUNCTION ─────────────────────────────────────
def convert_to_pdf(docx_path, pdf_path):
    try:
        word = comtypes.client.CreateObject("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(docx_path)
        doc.SaveAs(pdf_path, FileFormat=17)  # 17 = wdFormatPDF
        doc.Close(False)
        word.Quit()
        print(f"✅ PDF saved: {pdf_path}")
    except Exception as e:
        print(f"❌ PDF conversion failed for {docx_path}: {e}")

# ─── GOOGLE SHEETS AUTH ──────────────────────────────────────────
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("vdartintern-project1-285fc1b1d8bc.json", scope)
client = gspread.authorize(creds)
sheet = client.open("Data sheet").sheet1

expected_headers = [
    'No', 'Emp ID', 'Asset Code', 'Name', 'Reg No', 'Email', 'Aadhar', 'Phone', 'School / College',
    'College Location', 'Major', 'Year', 'Department', 'Role', 'D.O.B', 'Gender', 'FG', 'Method', 'Refered From',
    'Person Status', 'Shift Timing', 'Working Days', 'Shift', 'Join Date',
    'End Date', 'Model', 'Offer Letter', 'Offer Signed', 'Completion Letter'
]
data = sheet.get_all_records(expected_headers=expected_headers)
enrollment_dict = {row['Reg No']: row for row in data}

# ─── MAIN LOOP ───────────────────────────────────────────────────
internship_base_number = 100
i = 2
for reg_no, student in enrollment_dict.items():
    offer_status = str(student['Offer Letter']).strip().lower()
    if offer_status == "sended":
        i += 1
        continue

    try:
        # ── Dates and IDs ────────────────────────────────────────
        send_date = datetime.now().strftime("%d-%b-%Y")
        next_date = (datetime.now() + timedelta(days=1)).strftime("%d-%b-%Y")
        join_date_str = student.get('Join Date')
        end_date_str = student.get('End Date')
        internship_id = f"A0{internship_base_number}"

        # ── Context for Word Template ─────────────────────────────
        context = {
            "date": send_date,
            "name": student["Name"],
            "reg_no": student["Reg No"],
            "major": student["Major"],
            "college": student["School / College"],
            "role": student["Role"],
            "internship_id": internship_id,
            "start_date": join_date_str,
            "end_date": end_date_str,
            "next_date": next_date
        }

        # ── Save DOCX ─────────────────────────────────────────────
        doc = DocxTemplate(TEMPLATE_PATH)
        doc.render(context)
        safe_name = student["Name"].replace(" ", "_").replace("/", "_")
        docx_file = os.path.join(DOCX_OUTPUT, f"{safe_name}_Offer.docx")
        pdf_file = os.path.join(PDF_OUTPUT, f"{safe_name}_Offer.pdf")
        doc.save(docx_file)

        # ── Convert to PDF ────────────────────────────────────────
        convert_to_pdf(docx_file, pdf_file)

        # ── Send Email ────────────────────────────────────────────
        yag = yagmail.SMTP("vdartintern@gmail.com", "yqzs pjrd kmpt rcdt")
        subject = f"VDart Academy | OJT Offer Letter | {student['Name']}"
        body = f"""
Dear {student['Name']},

Congratulations on being selected for the OJT in {student['Role']} at VDart Academy.

Please find attached your onboarding letter. Kindly review, sign, and share it with us.

Welcome aboard!
"""
        yag.send(to=student["Email"], subject=subject, contents=[body, pdf_file])
        print(f"📧 Email sent to {student['Email']}")
        sheet.update_cell(i, 27, "Sended")

    except Exception as e:
        print(f"❌ Error for {student['Name']}: {e}")
        sheet.update_cell(i, 27, "Not Sended")

    internship_base_number += 1
    i += 1

print("🎉 All letters processed and emailed.")
