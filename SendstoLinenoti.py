import requests
import os
from pdf2docx import Converter

Line_Token = "ใส่ Token ที่ก็อปมาตรงนี้ได้เลย"
Line_Url = "https://notify-api.line.me/api/notify"
pdf_path = r"%CurrentFile.FullName%"
docx_path = os.path.splitext(pdf_path)[0] + ".docx"

try:
    cv = Converter(pdf_path)
    cv.convert(docx_path , start = 0 , end = None)
    cv.close()

    headers = {'Authorization': f'Bearer {Line_Token}'}
    message = f'แปลงไฟล์สำเร็จแล้ว \n ไฟล์: {os.path.basename(docx_path)}'

    payload = {'message': message}
    files = {'file': open(docx_path , 'rb')}
    response = requests.post(Line_Url , headers = headers , params = payload , files = files)

    if response.status_code == 200:
        print("ส่ง Line Notification สำเร็จ")
    else:
        print(f"ส่งไม่เป็นผลสำเร็จ: {response.status_code}")

except Exception as e:
    print(f"Error: {str(e)}")