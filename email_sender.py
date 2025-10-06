import os
import smtplib
import logging
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header
from dotenv import load_dotenv

# --- ЗАГРУЗКА ПЕРЕМЕННЫХ ИЗ .ENV ---
load_dotenv()
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
EMAIL_FROM = os.getenv("EMAIL_FROM")
EMAIL_TO = os.getenv("EMAIL_TO")

# --- ДАННЫЕ ДЛЯ ОТПРАВКИ EMAIL ---
SMTP_SERVER = 'smtp.mail.ru'
SMTP_PORT = 465


async def send_email_with_attachment(file_path):
    """Отправляет файл по email."""
    if not EMAIL_PASSWORD or not EMAIL_FROM or not EMAIL_TO:
        logging.error("Не найдены переменные окружения для отправки email. Убедитесь, что они есть в файле .env.")
        return False

    msg = MIMEMultipart()
    msg['From'] = EMAIL_FROM
    msg['To'] = EMAIL_TO
    msg['Subject'] = f"Отчёт по остаткам - {datetime.now().strftime('%d.%m.%Y')}"

    body = "Здравствуйте! Высылаю остатки.\n\nС уважением, Анна Сидорова."
    msg.attach(MIMEText(body, 'plain'))

    try:
        with open(file_path, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
        encoders.encode_base64(part)

        file_name = os.path.basename(file_path)
        part.add_header(
            "Content-Disposition",
            "attachment",
            filename=Header(file_name, 'utf-8').encode()
        )
        msg.attach(part)

        logging.info("Попытка установить соединение с SMTP-сервером...")
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(EMAIL_FROM, EMAIL_PASSWORD)
            server.sendmail(EMAIL_FROM, EMAIL_TO, msg.as_string())
        logging.info(f"Файл {os.path.basename(file_path)} успешно отправлен на почту {EMAIL_TO}")
        return True

    except smtplib.SMTPAuthenticationError:
        logging.error("Ошибка аутентификации SMTP. Проверьте логин/пароль приложения.")
        return False
    except Exception as e:
        logging.error(f"Не удалось отправить email: {e}", exc_info=True)
        return False
