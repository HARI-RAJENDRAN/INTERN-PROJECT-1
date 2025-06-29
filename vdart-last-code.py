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
from inference_sdk import InferenceHTTPClient
import cv2

# --- PATHS ---
TEMPLATE_PATH = r"C:\Users\tncha\Downloads\vdart-intern\pdf-template.docx"
DOCX_OUTPUT = r"C:\Users\tncha\Downloads\vdart-intern\formatted\docx"
PDF_OUTPUT = r"C:\Users\tncha\Downloads\vdart-intern\formatted\pdf"
EXTRACTED_PDF_DIR = "extracted_pdfs"
SIGNED_PDF_DIR = "signed_pdfs"
DETECTION_OUTPUT_DIR = "signature_detections"
POPPLER_PATH = r'C:\Users\tncha\Downloads\vdart-intern\poppler-24.08.0\Library\bin'

# Create required directories
for directory in [DOCX_OUTPUT, PDF_OUTPUT, EXTRACTED_PDF_DIR, SIGNED_PDF_DIR, DETECTION_OUTPUT_DIR]:
    os.makedirs(directory, exist_ok=True)

# --- PDF Conversion ---
def convert_to_pdf(docx_path, pdf_path):
    try:
        word = comtypes.client.CreateObject("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(docx_path)
        doc.SaveAs(pdf_path, FileFormat=17)
        doc.Close(False)
        word.Quit()
        print(f"PDF saved: {pdf_path}")
    except Exception as e:
        print(f"PDF conversion failed: {e}")

# --- Google Sheets Setup ---
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
        sheet.update_cell(1, len(header_row)+1, header_name)
        header_row.append(header_name)
        return len(header_row)

# Column indices
col_offer_letter = get_col_idx("Offer Letter")
col_offer_signed = get_col_idx("Offer Signed")
col_replied = get_col_idx("Replied")
col_pdf_link = get_col_idx("Extracted PDF Link")
col_sign_status = get_col_idx("Signature Verified")

# --- Roboflow Setup ---
ROBOFLOW_CLIENT = InferenceHTTPClient(
    api_url="https://detect.roboflow.com",
    api_key="xmudXzJexHGqhTJbyQx8"
)
MODEL_ID = "signature-krkm0/1"

def detect_signature_in_bottom_right(pdf_path):
    try:
        images = convert_from_path(pdf_path, first_page=1, last_page=1, poppler_path=POPPLER_PATH)
        if not images:
            return False

        image_path = pdf_path.replace('.pdf', '.jpg')
        images[0].save(image_path, "JPEG")

        result = ROBOFLOW_CLIENT.infer(image_path, model_id=MODEL_ID)

        image = cv2.imread(image_path)
        if image is None:
            return False

        height, width = image.shape[:2]
        signature_x_threshold = 0.65 * width
        signature_y_threshold = 0.80 * height

        valid_signatures = []
        for prediction in result.get("predictions", []):
            x, y, w, h = prediction["x"], prediction["y"], prediction["width"], prediction["height"]
            confidence = prediction.get("confidence", 0)
            if x > signature_x_threshold and y > signature_y_threshold:
                valid_signatures.append(prediction)
                x1, y1 = int(x - w/2), int(y - h/2)
                x2, y2 = int(x + w/2), int(y + h/2)
                cv2.rectangle(image, (x1, y1), (x2, y2), (0, 255, 0), 3)
                cv2.putText(image, f"Signature ({confidence:.2f})", (x1, y1-10), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 255, 0), 2)

        output_path = os.path.join(DETECTION_OUTPUT_DIR, os.path.basename(image_path))
        cv2.imwrite(output_path, image)
        print(f"Detection saved: {output_path}")

        return len(valid_signatures) > 0

    except Exception as e:
        print(f"Signature detection error: {e}")
        return False

# --- Email Sending ---
def send_offer_letters():
    yag = yagmail.SMTP("vdartintern@gmail.com", "yqzs pjrd kmpt rcdt")
    internship_base_number = 100
    i = 2
    for idx, row in enumerate(data, start=2):
        if str(row.get("Offer Letter", "")).strip().lower() == "sent":
            i += 1
            continue

        try:
            send_date = datetime.now().strftime("%d-%b-%Y")
            next_date = (datetime.now() + timedelta(days=1)).strftime("%d-%b-%Y")
            join_date_str = row.get('Join Date', '')
            end_date_str = row.get('End Date', '')
            internship_id = f"A0{internship_base_number}"

            context = {
                "date": send_date,
                "n_date": next_date,
                "name": row.get("Name", ""),
                "reg_no": row.get("Reg No", ""),
                "major": row.get("Major", ""),
                "college": row.get("School / College", ""),
                "role": row.get("Role", ""),
                "internship_id": internship_id,
                "start_date": join_date_str,
                "end_date": end_date_str,
                "next_date": next_date
            }

            doc = DocxTemplate(TEMPLATE_PATH)
            doc.render(context)
            safe_name = row.get("Name", "").replace(" ", "_").replace("/", "_")
            docx_file = os.path.join(DOCX_OUTPUT, f"{safe_name}_Offer.docx")
            pdf_file = os.path.join(PDF_OUTPUT, f"{safe_name}_Offer.pdf")
            doc.save(docx_file)

            convert_to_pdf(docx_file, pdf_file)

            email_id = str(row.get("Email", "")).strip().replace('`', '')
            subject = f"VDart Academy | On-the-Job Training – Onboarding Letter | {row.get('Name', '')}"
            body = f"""
            <div style="font-family:Arial, sans-serif; font-size:15px; color:#000;">
                <p><strong>Dear {row.get('Name', '')},</strong></p>
                Congratulations on being selected for the <strong>On-the-Job Training (OJT)</strong> in 
                <strong>{row.get('Role', '')}</strong> with <strong>VDart Academy</strong>.<br><br>
                We are pleased to inform you that your OJT will commence on 
                <strong>{join_date_str}</strong>.<br><br>
                Please find your onboarding letter attached. Kindly review, countersign, and share it with us to confirm your acceptance.
                We look forward to a meaningful and successful journey ahead. If you have any questions, feel free to reach out.
                <p><strong>Welcome aboard!</strong></p>
            </div>
            """

            yag.send(to=email_id, subject=subject, contents=[body, pdf_file])
            print(f"Email sent to {email_id}")
            sheet.update_cell(i, col_offer_letter, "Sent")

        except Exception as e:
            print(f"Failed to send email to {row.get('Email', '')}: {e}")
            sheet.update_cell(i, col_offer_letter, "Not Sent")

        internship_base_number += 1
        i += 1

# --- Email Parsing ---
def process_replied_emails():
    print("Checking for replied emails...")
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login("vdartintern@gmail.com", "yqzs pjrd kmpt rcdt")
        mail.select("inbox")

        status, messages = mail.search(None, 'ALL')
        email_attachments = {}

        if status == "OK":
            for num in messages[0].split():
                typ, msg_data = mail.fetch(num, "(RFC822)")
                for response_part in msg_data:
                    if isinstance(response_part, tuple):
                        msg = email.message_from_bytes(response_part[1])
                        sender = msg.get("From")
                        match = re.search(r'<(.+?)>', sender)
                        email_address = match.group(1) if match else sender
                        email_address = email_address.strip().lower()

                        pdf_paths = []
                        for part in msg.walk():
                            if part.get_content_maintype() == 'multipart':
                                continue
                            filename = part.get_filename()
                            content_type = part.get_content_type()
                            if (content_type == 'application/pdf' or 
                                (filename and filename.lower().endswith('.pdf'))):
                                if not filename:
                                    filename = f"{email_address}_reply.pdf"
                                filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
                                save_path = os.path.join(EXTRACTED_PDF_DIR, filename)
                                with open(save_path, 'wb') as f:
                                    f.write(part.get_payload(decode=True))
                                pdf_paths.append(save_path)
                        if pdf_paths:
                            email_attachments[email_address] = pdf_paths

        mail.logout()
        return email_attachments

    except Exception as e:
        print(f"Email fetch error: {e}")
        return {}

# --- Process Signed PDFs and Update Sheet ---
def process_signed_pdfs():
    replied_emails = process_replied_emails()

    for idx, row in enumerate(data, start=2):
        email_id = str(row.get("Email", "")).strip().lower().replace('`', '')
        if email_id in replied_emails:
            for pdf_path in replied_emails[email_id]:
                signed = detect_signature_in_bottom_right(pdf_path)
                sheet.update_cell(idx, col_replied, "Yes")
                sheet.update_cell(idx, col_pdf_link, pdf_path)
                if signed:
                    signed_pdf_path = os.path.join(SIGNED_PDF_DIR, os.path.basename(pdf_path))
                    shutil.copy(pdf_path, signed_pdf_path)
                    print(f"Signature Verified: {signed_pdf_path}")
                    sheet.update_cell(idx, col_sign_status, "Signature Verified")
                    sheet.update_cell(idx, col_offer_signed, "Signed")
                else:
                    print("No signature detected.")
                    sheet.update_cell(idx, col_sign_status, "No Valid Signature")
                    sheet.update_cell(idx, col_offer_signed, "Not Signed")

# --- MAIN ---
if __name__ == "__main__":
    print("Starting signature verification process...")
    send_offer_letters()
    process_signed_pdfs()
    print("Done. Google Sheet updated.")