# Author: Apiwat Kongsawat (Jimmy)
# Modify: Maimongkol Thokanokwan (Woody)

from dataclasses import dataclass
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
import os
from dotenv import load_dotenv

load_dotenv(dotenv_path=os.path.join(os.path.dirname(__file__), '..', '.env'))

@dataclass
class EmailConfig:
    sender_email: str
    sender_password: str
    smtp_server: str
    smtp_port: int
    recipient_csv_path: str
    recipient_attachment_folder: str
    email_subject : str


def send_email(message, recipient, attachment_path, email_config: EmailConfig) -> bool:
    try:
        msg = MIMEMultipart()
        msg['From'] = email_config.sender_email
        msg['To'] = recipient
        msg['Subject'] = email_config.email_subject

        msg.attach(MIMEText(message, 'plain'))
        # Attach the file
        attachment = open(attachment_path, "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename= {os.path.basename(attachment_path)}')
        msg.attach(part)
        
        server = smtplib.SMTP(email_config.smtp_server, email_config.smtp_port)
        server.starttls()
        server.login(email_config.sender_email, email_config.sender_password)
        server.send_message(msg)
        server.quit()
        return True
        
    except Exception as e:
        print(f"Failed to send email to {recipient}. Error: {e}")
        return False

def load_data_from_excel(file_path :str , sheet_name = 0):
    try:
        # Load student scores from the first sheet
        df_scores = pd.read_excel(file_path, sheet_name=sheet_name) # เปลี่ยน sheet name (index เริ่ม 0) เพื่อเปลี่ยนชีท
        
        # Open the workbook to access specific cells for email configuration
        workbook = pd.ExcelFile(file_path)

        return df_scores

    except Exception as e:
        print(f"Error loading Excel data: {e}")
        return None, None

def get_certificate_path(student, folder_path):
    try:
        filename = f"{folder_path}/cer_{student['student_examid']}.txt" 
        if os.path.exists(filename):
            return filename
        else:
            print(f"Certificate file for student {student['student_id']} does not exist.")
            return None
    except Exception as e:
        print(f"Failed to get certificate file for student {student['student_id']}. Error: {e}")
        return None

def load_EmailConfig() -> EmailConfig:
    # Load environment variables and validate presence
    required_vars = [
        "SENDER_EMAIL", "SENDER_PASSWORD",
        "SMTP_SERVER", "SMTP_PORT",
        "RECIPIENT_CSV_PATH", "RECIPIENT_ATTACHMENT_FOLDER", "EMAIL_SUBJECT"
    ]

    missing = [var for var in required_vars if not os.getenv(var)]
    if missing:
        raise EnvironmentError(f"Missing environment variables: {', '.join(missing)}")

    return EmailConfig(
        sender_email=os.getenv("SENDER_EMAIL"),
        sender_password=os.getenv("SENDER_PASSWORD"),
        smtp_server=os.getenv("SMTP_SERVER"),
        smtp_port=int(os.getenv("SMTP_PORT")),
        recipient_csv_path=os.getenv("RECIPIENT_CSV_PATH"),
        recipient_attachment_folder=os.getenv("RECIPIENT_ATTACHMENT_FOLDER"),
        email_subject = os.getenv("EMAIL_SUBJECT")
    )

def main():
    try:
        file_path = os.getenv('RECIPIENT_CSV_PATH')
        folder_path = os.getenv('RECIPIENT_ATTACHMENT_FOLDER')

        # Load email config from .env
        email_config = load_EmailConfig()

        # Load student data and email configurations from Excel
        data = load_data_from_excel(file_path, 'Sheet1')

        if data is not None:
            for index, student in data.iterrows():
                if not pd.notna(student['email']) or str(student['email']).strip() == "": 
                    print(f"Skipping row {index+1}: No email address")
                    continue
                personalized_info = f"\n\nรหัสประจำตัวของผู้เข้าสอบ: {student['student_examid']}\nชื่อผู้เข้าสอบ: {student['name']} {student['surname']}\nโรงเรียน {student['school']}\nห้องสอบ: {student['room']}\nเวลาสอบ: {student['time']}"
                message = f"ถึง {student['name']} {student['surname']}\nตามที่นักเรียนได้สมัครเข้าร่วมการแข่งขัน MU Mental Math Competition 2025\n\n"  + \
                "นักเรียนสามารถดาวน์โหลดบัตรประจำตัวผู้เข้าสอบที่แนบมากับอีเมลนี้ได้เลยนะครับ " + personalized_info\
                + "\n" + "\n" + "อย่าลืมนำบัตรประจำตัวผู้เข้าสอบและบัตรประชาชนมาเพื่อใช้ในการยืนยันตัวตน และเช็คความถูกต้องของข้อมูลก่อนเข้าห้องสอบด้วยครับ" + "\n\n" + \
                "อีเมลนี้ถูกสร้างขึ้นโดยอัตโนมัติ หากมีข้อสงสัยเพิ่มเติม สามารถสอบถาม[ชื่อคนให้ติดต่อ] \n" + "ได้ที่อีเมล: email@gmail.com, MoreEmail@gmail \nและเบอร์โทร 000-000-0000\n" \
                + "\n -- \n\n[Firstname] [Surname], Student at Mahidol University\nDepartment of [Department name], Mahidol University\nRama VI Road, Ratchathewi, Bangkok 10400, Thailand\nEmail: UniversityMail@mahidol.edu\nTel: +66x xxxx xxxx"
                
                attachment_path = get_certificate_path(student, folder_path)
                if attachment_path:
                    if(send_email(message, student['email'], attachment_path, email_config)):
                        print(f"Email successfully sent to {student['email']} (ID: {index+1})")
                    else:
                        print(f"Can't sent email to {student['email']} (ID: {index+1})")
                    

    except FileNotFoundError as e:
        print(f"File not found: {e}")
    except EnvironmentError as e:
        print(f"Environment setup issue: {e}")
    except Exception as e:
        print(f"Unexpected error: {e}")

        
if __name__ == "__main__":
    main()
