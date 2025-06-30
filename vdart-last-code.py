import gspread
from oauth2client.service_account import ServiceAccountCredentials
import yagmail
from datetime import datetime, timedelta
import os
import comtypes.client
import imaplib
import email
import re
import shutil
from docxtpl import DocxTemplate
from pdf2image import convert_from_path
import cv2
import numpy as np
import getpass

# --- PATHS ---
TEMPLATE_PATH = r"C:\Users\academytraining\Documents\HARI-RAJENDRAN\pdf-template.docx"
DOCX_OUTPUT = r"C:\Users\academytraining\Documents\HARI-RAJENDRAN\formatted\docx"
PDF_OUTPUT = r"C:\Users\academytraining\Documents\HARI-RAJENDRAN\formatted\pdf"
EXTRACTED_PDF_DIR = "extracted_pdfs"
SIGNED_PDF_DIR = "signed_pdfs"
DETECTION_OUTPUT_DIR = "signature_detections"
POPLER_PATH = r"C:\Users\academytraining\Documents\HARI-RAJENDRAN\Release-24.08.0-0\poppler-24.08.0\Library\bin"

# --- MAKE SURE DIRECTORIES EXIST ---
os.makedirs(DOCX_OUTPUT, exist_ok=True)
os.makedirs(PDF_OUTPUT, exist_ok=True)
os.makedirs(EXTRACTED_PDF_DIR, exist_ok=True)
os.makedirs(SIGNED_PDF_DIR, exist_ok=True)
os.makedirs(DETECTION_OUTPUT_DIR, exist_ok=True)

# --- PDF CONVERSION ---
def convert_to_pdf(docx_path, pdf_path):
    try:
        word = comtypes.client.CreateObject("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(docx_path)
        doc.SaveAs(pdf_path, FileFormat=17)
        doc.Close(False)
        word.Quit()
        print(f"[✓] PDF saved: {pdf_path}")
    except Exception as e:
        print(f"[x] PDF conversion failed: {e}")

# --- GOOGLE SHEETS SETUP ---
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("vdartintern-project1-285fc1b1d8bc.json", scope)
client = gspread.authorize(creds)
sheet = client.open("Data sheet").sheet1
data = sheet.get_all_records()
header_row = sheet.row_values(1)

def get_col_idx(header_name):
    try:
        return header_row.index(header_name) + 1
    except ValueError:
        sheet.update_cell(1, len(header_row) + 1, header_name)
        header_row.append(header_name)
        return len(header_row)

col_offer_letter = get_col_idx("Offer Letter")
col_offer_signed = get_col_idx("Offer Signed")
col_replied = get_col_idx("Replied")
col_pdf_link = get_col_idx("Extracted PDF Link")
col_sign_status = get_col_idx("Signature Verified")

# --- SIGNATURE DETECTION ---
def is_signature_present(pdf_path):
    try:
        images = convert_from_path(pdf_path, first_page=1, last_page=1, poppler_path=POPLER_PATH)
        if not images:
            return False

        image_path = pdf_path.replace(".pdf", ".jpg")
        images[0].save(image_path, "JPEG")
        image = cv2.imread(image_path)
        if image is None:
            return False

        height, width = image.shape[:2]
        x_start = int(width * 0.75)
        y_start = int(height * 0.75)
        cropped = image[y_start:, x_start:]

        gray = cv2.cvtColor(cropped, cv2.COLOR_BGR2GRAY)
        _, thresh = cv2.threshold(gray, 180, 255, cv2.THRESH_BINARY_INV)
        contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        total_area = cropped.shape[0] * cropped.shape[1]
        ink_area = sum(cv2.contourArea(c) for c in contours if cv2.contourArea(c) > 150)
        ink_ratio = ink_area / total_area

        vis_path = os.path.join(DETECTION_OUTPUT_DIR, os.path.basename(image_path).replace(".jpg", "_detected.jpg"))
        cv2.rectangle(image, (x_start, y_start), (width, height), (0, 255, 0), 2)
        cv2.imwrite(vis_path, image)
        print(f"[🖼️] Detection preview saved: {vis_path}")

        return ink_ratio > 0.01
    except Exception as e:
        print(f"[x] Signature detection error: {e}")
        return False

# --- SENDING OFFER LETTERS ---
print("[🚀] Starting offer letter generation and sending...")
internship_base_number = 100
i = 2

yag = yagmail.SMTP("vdartintern@gmail.com", "yqzs pjrd kmpt rcdt")

for row in data:
    offer_status = str(row.get("Offer Letter", "")).strip().lower()
    if offer_status == "sended":
        i += 1
        continue

    try:
        send_date = datetime.now().strftime("%d-%b-%Y")
        next_date = (datetime.now() + timedelta(days=1)).strftime("%d-%b-%Y")
        internship_id = f"A0{internship_base_number}"
        context = {
            "date": send_date,
            "n_date": next_date,
            "name": row["Name"],
            "reg_no": row["Reg No"],
            "major": row.get("Major", ""),
            "college": row.get("School / College", ""),
            "role": row.get("Role", ""),
            "internship_id": internship_id,
            "start_date": row.get("Join Date", ""),
            "end_date": row.get("End Date", "")
        }

        doc = DocxTemplate(TEMPLATE_PATH)
        doc.render(context)
        safe_name = re.sub(r"[^\w\-]", "_", row["Name"])
        docx_file = os.path.join(DOCX_OUTPUT, f"{safe_name}_Offer.docx")
        pdf_file = os.path.join(PDF_OUTPUT, f"{safe_name}_Offer.pdf")
        doc.save(docx_file)
        convert_to_pdf(docx_file, pdf_file)

        subject = f"VDart Academy | On-the-Job Training – Onboarding Letter | {row['Name']}"
        body = f"""
            <div style="font-family:Arial, sans-serif; font-size:15px; color:#fff;">
                <p><strong>Dear {row.get('Name', '')},</strong></p>
                Congratulations on being selected for the <strong>On-the-Job Training (OJT)</strong> in 
                <strong>{row.get('Role', '')}</strong> with <strong>VDart Academy</strong>.<br>
                We are pleased to inform you that your OJT will commence on 
                <strong>{row.get('Join Date', '')}</strong>.<br>
                Please find your onboarding letter attached. Kindly review, countersign, and share it with us to confirm your acceptance.
                <p><strong>Welcome aboard!</strong></p>
            </div>
        """
        yag.send(to=row["Email"], subject=subject, contents=[body, pdf_file])
        print(f"[📧] Sent email to {row['Email']}")
        sheet.update_cell(i, col_offer_letter, "Sended")

    except Exception as e:
        print(f"[x] Error for {row['Name']}: {e}")
        sheet.update_cell(i, col_offer_letter, "Not Sended")

    internship_base_number += 1
    i += 1

print("[✅] All letters processed and mailed.")

# --- PROCESS REPLIES ---
def check_replied_emails_and_process_pdfs(sheet, data):
    print("[🔍] Checking for replied emails with PDFs...")
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login()
        mail.select("inbox")
        status, messages = mail.search(None, "ALL")
        replied_mail = dict()

        for num in messages[0].split():
            try:
                typ, msg_data = mail.fetch(num, "(RFC822)")
                for response_part in msg_data:
                    if isinstance(response_part, tuple):
                        msg = email.message_from_bytes(response_part[1])
                        sender = msg.get("From", "")
                        match = re.search(r"<(.+?)>", sender)
                        email_address = match.group(1) if match else sender
                        email_address = email_address.strip().lower()
                        pdf_paths = replied_mail.get(email_address, [])
                        found_pdf = False

                        for part in msg.walk():
                            if part.get_content_maintype() == "multipart":
                                continue
                            filename = part.get_filename()
                            content_type = part.get_content_type()
                            content_disposition = str(part.get("Content-Disposition", ""))

                            is_pdf = (
                                content_type == "application/pdf"
                                or (filename and filename.lower().endswith(".pdf"))
                                or "attachment" in content_disposition.lower()
                            )
                            if is_pdf:
                                if not filename:
                                    filename = f"{email_address}_extracted_{len(pdf_paths)+1}.pdf"
                                filename = re.sub(r"[<>:\"/\\|?*]", "_", filename)
                                save_path = os.path.join(EXTRACTED_PDF_DIR, filename)
                                with open(save_path, "wb") as f:
                                    f.write(part.get_payload(decode=True))
                                print(f"[📎] Saved PDF: {save_path}")
                                pdf_paths.append(save_path)
                                found_pdf = True

                        if found_pdf:
                            replied_mail[email_address] = pdf_paths
            except Exception as e:
                print(f"[x] Email processing error: {e}")
                continue

        mail.logout()

        for idx, row in enumerate(data, start=2):
            email_id = str(row.get("Email", "")).strip().lower()
            if email_id in replied_mail:
                print(f"[✔] {email_id} replied. Checking signatures...")
                for local_pdf_path in replied_mail[email_id]:
                    extracted_pdf_link = local_pdf_path
                    signature_verified = is_signature_present(local_pdf_path)
                    if signature_verified:
                        signed_pdf_path = os.path.join(SIGNED_PDF_DIR, os.path.basename(local_pdf_path))
                        shutil.copy(local_pdf_path, signed_pdf_path)
                        print(f"[✓] Signature detected: {signed_pdf_path}")
                        signature_status = "Signature Verified"
                        offer_signed = "Signed"
                    else:
                        print(f"[x] No signature: {local_pdf_path}")
                        signature_status = "No Signature Detected"
                        offer_signed = "Not Signed"

                    sheet.update_cell(idx, col_replied, "Yes")
                    sheet.update_cell(idx, col_pdf_link, extracted_pdf_link)
                    sheet.update_cell(idx, col_sign_status, signature_status)
                    sheet.update_cell(idx, col_offer_signed, offer_signed)
            else:
                print(f"[ℹ️] No PDF reply from: {email_id}")
    except Exception as e:
        print(f"[x] Email check failed: {e}")

# --- RUN CHECK ---
check_replied_emails_and_process_pdfs(sheet, data)
print("[🎯] Completed checking replies, detecting signatures, and updating sheet.")

