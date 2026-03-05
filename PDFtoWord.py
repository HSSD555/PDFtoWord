import os
pdf_path = r"%CurrentFile.FullName%"
docx_path = os.path.splitext(pdf_path)[0] + ".docx"

try:
    cv = Converter(pdf_path)
    cv.convert(docx_path , start = 0 , end = None)
    cv.close()
    print(f"Success: แปลงเป็น Word เรียบร้อยที่ {docx_path}")

except Exception as e:
    print(f"Error: {str(e)}")

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

sender_email = "your_email@gmail.com"
receiver_email = "target_email@gmail.com"
password = "your_app_password_here"
file_to_attach = docx_path

msg = MIMEMultipart()
msg['From'] = sender_email
msg['Subject'] = f"ส่งไฟล์ที่แปลงแล้ว: {os.path.basename(file_to_attach)}"

body = "แปลงไฟล์เรียบร้อยแล้ว รายละเอียดดังนี้"
msg.attach(MIMEText(body , 'plain'))

try:
    with open(file_to_attach , "rb") as attachment:
        part = MIMEBase("application" , "octet-stream")
        part.set_payload(attach.read())

    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        f"attachment; filename = {os.path.basename(file_to_attach)}",
    )

    msg.attach(part)
    server = smtplib.SMTP('stmp.gmail.com' , 587)
    server.starttls()
    server.login(sender_email , password)
    server.send_mail(sender_email , receiver_email , msg.as_string())
    server.quit()

    print("Email Sent Successfully!")

except Exception as e:
    print(f"Failed to send email: {str(e)}")