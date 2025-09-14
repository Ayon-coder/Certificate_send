import os
import pdfkit
import pandas as pd
from datetime import datetime
from dotenv import load_dotenv
import smtplib
from email.message import EmailMessage

# 🔹 Load .env
load_dotenv("email.env")

SMTP_HOST = os.getenv("SMTP_HOST")
SMTP_PORT = int(os.getenv("SMTP_PORT", 465))
SMTP_USER = os.getenv("SMTP_USER")
SMTP_PASS = os.getenv("SMTP_PASS")

if not SMTP_USER or not SMTP_PASS:
    raise ValueError("Set SMTP_USER and SMTP_PASS in .env")

# 🔹 Path to wkhtmltopdf (update if installed elsewhere)
WKHTMLTOPDF_PATH = r"file_path_of\wkhtmltopdf.exe"
config = pdfkit.configuration(wkhtmltopdf=WKHTMLTOPDF_PATH)

# 🔹 PDF options (allow local file access)
options = {
    "enable-local-file-access": ""
}

# 🔹 Load certificate template
with open("certificate.html", "r", encoding="utf-8") as f:
    template = f.read()

# 🔹 Load student list
students = pd.read_csv("students.csv")

results = []

for _, row in students.iterrows():
    name, email = row["name"], row["email"]
    date_str = datetime.today().strftime("%d %B %Y")

    # Fill template
    html_content = template.replace("{{name}}", name).replace("{{date}}", date_str)

    # Save PDF
    pdf_file = f"certificates/{name.replace(' ', '_')}.pdf"
    os.makedirs("certificates", exist_ok=True)
    pdfkit.from_string(html_content, pdf_file, configuration=config, options=options)

    # Create email
    msg = EmailMessage()
    msg["Subject"] = "Your Certificate"
    msg["From"] = SMTP_USER
    msg["To"] = email
    msg.set_content(f"Dear {name},\n\nPlease find attached your certificate.\n\nRegards,\nOrganizer")

    with open(pdf_file, "rb") as f:
        msg.add_attachment(f.read(), maintype="application", subtype="pdf", filename=os.path.basename(pdf_file))

    try:
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as server:
            server.login(SMTP_USER, SMTP_PASS)
            server.send_message(msg)
        status = "Sent"
    except Exception as e:
        status = f"Failed: {e}"

    print(f"{name} -> {status}")
    results.append({"name": name, "email": email, "status": status})

# 🔹 Save results
pd.DataFrame(results).to_csv("result.csv", index=False)
print("\n✅ All done! Check result.csv for the summary.")

