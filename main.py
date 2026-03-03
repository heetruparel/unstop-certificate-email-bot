import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import yagmail
import os
import sys
import time
from dotenv import load_dotenv

# ==============================
# LOAD EMAIL CREDENTIALS
# ==============================

load_dotenv()
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
APP_PASSWORD = os.getenv("APP_PASSWORD")

if not SENDER_EMAIL or not APP_PASSWORD:
    print("Email credentials missing in .env file")
    sys.exit(1)

# ==============================
# CHECK EVENT ARGUMENT
# ==============================

if len(sys.argv) != 2:
    print("Usage: python3 main.py event_name")
    sys.exit(1)

event_name = sys.argv[1]

event_path = f"events/{event_name}"
csv_path = f"{event_path}/data.csv"
template_path = f"{event_path}/template.png"

if not os.path.exists(csv_path):
    print("CSV file not found!")
    sys.exit(1)

if not os.path.exists(template_path):
    print("Template file not found!")
    sys.exit(1)

# ==============================
# CREATE OUTPUT FOLDER
# ==============================

output_folder = f"generated/{event_name}"
os.makedirs(output_folder, exist_ok=True)

# ==============================
# READ CSV (UNSTOP FORMAT)
# ==============================

data = pd.read_csv(csv_path)

name_column = None
email_column = None
payment_column = None

for col in data.columns:
    lower = col.lower()
    if "name" in lower:
        name_column = col
    if "email" in lower:
        email_column = col
    if "payment" in lower or "status" in lower:
        payment_column = col

if not name_column or not email_column:
    print("Could not detect Name or Email column.")
    sys.exit(1)

# ==============================
# FILTER PAID PARTICIPANTS ONLY
# ==============================

if payment_column:
    data = data[
        data[payment_column]
        .astype(str)
        .str.lower()
        .isin(["paid", "successful", "completed"])
    ]
    print(f"Using payment column: {payment_column}")
else:
    print("Payment column not detected. Sending to all.")

data = data.drop_duplicates(subset=[email_column])

print(f"Total certificates to send: {len(data)}")

# ==============================
# EMAIL SETUP
# ==============================

yag = yagmail.SMTP(SENDER_EMAIL, APP_PASSWORD)

# ==============================
# START GENERATING CERTIFICATES
# ==============================

for index, row in data.iterrows():

    name = str(row[name_column]).strip()
    email = str(row[email_column]).strip()

    image = Image.open(template_path)
    draw = ImageDraw.Draw(image)

    # ===== PROFESSIONAL FONT SETTINGS =====
    font = ImageFont.truetype(
        "/System/Library/Fonts/Supplemental/Georgia.ttf", 65
    )

    width, height = image.size

    # Get text bounding box
    bbox = draw.textbbox((0, 0), name, font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]

    # Center horizontally
    x = (width - text_width) / 2 +80  # Slight right shift for better alignment

    # Move name slightly upward for proper alignment
    y = height * 0.49

    # Draw name (dark grey for premium look)
    draw.text((x, y), name, fill="#222222", font=font)

    # Save as PDF
    file_path = f"{output_folder}/{name}.pdf"
    image.save(file_path)

    # ==============================
    # PROFESSIONAL EMAIL BODY
    # ==============================

    email_body = f"""
Dear {name},

Thank you for attending VEGA Hackathon.

We sincerely appreciate your participation and the enthusiasm you brought to the event.

Attached is your Certificate of Participation.

We look forward to welcoming you to our future events.

Warm regards,
Organizing Team
"""

    yag.send(
        to=email,
        subject="Certificate - VEGA Hackathon",
        contents=email_body,
        attachments=file_path
    )

    print(f"Sent to {name}")
    time.sleep(2)  # Prevent Gmail blocking

print("All paid participant certificates sent successfully!")