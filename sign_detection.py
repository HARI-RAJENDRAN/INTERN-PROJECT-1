from inference_sdk import InferenceHTTPClient
import cv2

# ─── CONFIGURATION ───────────────────────────────────────────────
API_KEY = "xmudXzJexHGqhTJbyQx8"  # Your Roboflow API key
MODEL_ID = "signature-krkm0/1"    # Model name/version
IMAGE_PATH = r"C:\Users\academytraining\Downloads\HARI-INTERN\VS\page_0.jpg"

# ─── 1. Connect to Roboflow API ─────────────────────────────────
CLIENT = InferenceHTTPClient(
    api_url="https://detect.roboflow.com",
    api_key=API_KEY
)

# ─── 2. Perform Inference ───────────────────────────────────────
result = CLIENT.infer(IMAGE_PATH, model_id=MODEL_ID)

print("✅ Detection complete!")
print(result)

# ─── 3. Draw Bounding Boxes ─────────────────────────────────────
image = cv2.imread(IMAGE_PATH)

for prediction in result["predictions"]:
    x, y, w, h = prediction["x"], prediction["y"], prediction["width"], prediction["height"]
    x1 = int(x - w / 2)
    y1 = int(y - h / 2)
    x2 = int(x + w / 2)
    y2 = int(y + h / 2)

    # Draw rectangle
    cv2.rectangle(image, (x1, y1), (x2, y2), (0, 255, 0), 2)

    # Draw label
    label = prediction["class"]
    confidence = prediction["confidence"]
    text = f"{label} ({confidence:.2f})"
    cv2.putText(image, text, (x1, y1 - 10),
                cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 255, 0), 2)

# ─── 4. Save the Output Image ───────────────────────────────────
output_path = "detected.jpg"
cv2.imwrite(output_path, image)
print(f"🖼️ Saved result to: {output_path}")
