import imaplib
import email
from email.header import decode_header
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from inference_sdk import InferenceHTTPClient
import mimetypes
from pdf2image import convert_from_path
import tempfile

# ─── CONFIGURATION ──────────────────────────────────────────────────────
EMAIL = "vdartintern@gmail.com"
PASSWORD = "yqzs pjrd kmpt rcdt"  # App password
IMAP_SERVER = "imap.gmail.com"
SAVE_FOLDER = r"C:\Users\Yogesh\Downloads\HARI_CHARLEZ\Sign-Detection-Process\email_replies"
PDF_IMAGES_FOLDER = os.path.join(SAVE_FOLDER, "converted_images")
images = convert_from_path(filepath, poppler_path=r"C:\Users\Yogesh\Downloads\Release-24.08.0-0 (1)\poppler-24.08.0\Library\bin")

os.makedirs(SAVE_FOLDER, exist_ok=True)
os.makedirs(PDF_IMAGES_FOLDER, exist_ok=True)

# ─── SIGNATURE DETECTION SETUP ───────────────────────────────────
API_KEY = "xmudXzJexHGqhTJbyQx8"
MODEL_ID = "signature-krkm0/1"
CLIENT = InferenceHTTPClient(api_url="https://detect.roboflow.com", api_key=API_KEY)

# ─── DECODE HELPERS ─────────────────────────────────
def decode_mime_words(s):
    return ''.join(
        word.decode(enc or 'utf-8') if isinstance(word, bytes) else word
        for word, enc in decode_header(s)
    )

def decode_filename(name):
    if not name:
        return None
    decoded, enc = decode_header(name)[0]
    if isinstance(decoded, bytes):
        return decoded.decode(enc or "utf-8", errors="ignore")
    return decoded

# ─── GOOGLE SHEET READ EMAILS ────────────────────
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("vdartintern-project1-285fc1b1d8bc.json", scope)
client = gspread.authorize(creds)
sheet = client.open("Data sheet").sheet1
data = sheet.get_all_records()
candidate_emails = {str(row.get("Email", "")).strip().lower(): idx+2 for idx, row in enumerate(data) if row.get("Email")}

# ─── CONNECT TO GMAIL ───────────────────────
try:
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL, PASSWORD)
    mail.select("INBOX")
except Exception as e:
    print(f"❌ Connection/Login failed: {e}")
    exit()

# ─── FETCH ALL EMAILS ───────────────────
status, messages = mail.search(None, 'ALL')
if status != "OK":
    print("❌ Could not fetch emails.")
    mail.logout()
    exit()

email_ids = messages[0].split()
print(f"📧 Found {len(email_ids)} emails.")

# ─── PROCESS EACH EMAIL ───────────────────
for email_id in email_ids:
    _, msg_data = mail.fetch(email_id, "(RFC822)")
    for resp in msg_data:
        if not isinstance(resp, tuple):
            continue

        msg = email.message_from_bytes(resp[1])
        subject = decode_mime_words(msg.get("Subject", ""))
        from_email = decode_mime_words(msg.get("From", ""))
        from_email_clean = from_email.split()[-1].strip("<>").lower()
        in_reply_to = msg.get("In-Reply-To", "")
        references = msg.get("References", "")

        is_reply = subject.lower().startswith("re:") or in_reply_to or references
        if not is_reply or from_email_clean not in candidate_emails:
            continue

        print(f"\n📨 Reply Email #{email_id.decode()}")
        print(f"   From: {from_email}")
        print(f"   Subject: {subject}")

        attachment_found = False
        part_counter = 1

        for part in msg.walk():
            content_type = part.get_content_type()
            filename = decode_filename(part.get_filename())
            content_disposition = str(part.get("Content-Disposition") or "")

            print("\n--- Email Part ---")
            print(f"Content-Type     : {content_type}")
            print(f"Disposition      : {content_disposition}")
            print(f"Filename         : {filename}")
            print("------------------------------")

            if part.get_content_maintype() == 'multipart':
                continue

            if not filename:
                ext = mimetypes.guess_extension(content_type.split(';')[0].strip()) or '.bin'
                filename = f"email_{email_id.decode()}_part{part_counter}{ext}"
                part_counter += 1

            filepath = os.path.join(SAVE_FOLDER, filename)

            try:
                payload = part.get_payload(decode=True)
                if payload and len(payload) > 512:
                    with open(filepath, "wb") as f:
                        f.write(payload)
                    print(f"✅ Saved: {filepath}")
                    attachment_found = True

                    row_index = candidate_emails[from_email_clean]

                    # ─── PDF to Image and Signature Detection ─────
                    if filename.lower().endswith(".pdf"):
                        try:
                            images = convert_from_path(filepath)
                            for i, image in enumerate(images):
                                img_path = os.path.join(PDF_IMAGES_FOLDER, f"{os.path.splitext(filename)[0]}_page{i+1}.jpg")
                                image.save(img_path, "JPEG")
                                result = CLIENT.infer(img_path, model_id=MODEL_ID)
                                print("🔍 Signature Detection Result:", result)
                                signed = any(pred.get("class", "").lower() in ["signature", "signature-detection"] for pred in result.get("predictions", []))
                                if signed:
                                    sheet.update_cell(row_index, 28, "Signed")
                                    break
                            else:
                                sheet.update_cell(row_index, 28, "Not Signed")
                        except Exception as e:
                            print(f"❌ Error processing PDF {filename}: {e}")
                    else:
                        # ─── Handle image directly ─────
                        if filename.lower().endswith(('.jpg', '.jpeg', '.png')):
                            result = CLIENT.infer(filepath, model_id=MODEL_ID)
                            print("🔍 Signature Detection Result:", result)
                            signed = any(pred.get("class", "").lower() in ["signature", "signature-detection"] for pred in result.get("predictions", []))
                            status = "Signed" if signed else "Not Signed"
                            sheet.update_cell(row_index, 28, status)

                else:
                    print("⚠️ Skipped part (empty or very small content).")
            except Exception as e:
                print(f"❌ Error saving {filename}: {e}")

        if not attachment_found:
            print("⚠️ No attachment found in this reply.")
            print("ℹ️ Reason: Possibly embedded inline or missing filename/disposition.")
            print("💡 Suggestion: Ask the sender to send the file as an actual attachment.")

# ─── CLEAN UP ──────────────────────
mail.logout()
print("\n📦 Done.")
