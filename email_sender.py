import smtplib
from email.message import EmailMessage
import mimetypes
import os

EMAIL_ADDRESS = "dlyashkolisusu@gmail.com"
EMAIL_PASSWORD = "wubfansclyihcymz"

HTML_FRAME = """\
<html>
  <body style="margin:0; padding:0; font-family: Arial, sans-serif;">
    <div style="background-color: orange; height: 8px; width: 100%;"></div>
    <div style="padding: 20px; text-align: left;">
      {content}
    </div>
  </body>
</html>
"""

def send_email_with_attachment(recipient_email, subject, content_html, attachments=None):
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = recipient_email

    msg.set_content("Это письмо содержит HTML-контент. Включите отображение HTML в почтовом клиенте.")
    msg.add_alternative(HTML_FRAME.format(content=content_html), subtype='html')

    # Добавление вложений
    if attachments:
        for file_path in attachments:
            if not os.path.isfile(file_path):
                print(f"Файл не найден: {file_path}")
                continue

            mime_type, _ = mimetypes.guess_type(file_path)
            if mime_type is None:
                mime_type = 'application/octet-stream'
            maintype, subtype = mime_type.split('/', 1)

            with open(file_path, 'rb') as f:
                file_data = f.read()
                file_name = os.path.basename(file_path)

            msg.add_attachment(file_data, maintype=maintype, subtype=subtype, filename=file_name)

    # Отправка
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        smtp.send_message(msg)

    print(f"Письмо с вложением отправлено на {recipient_email}")

def send_email(recipient_email, subject, content_html):
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = recipient_email

    msg.set_content("Это письмо содержит HTML-контент. Включите отображение HTML в почтовом клиенте.")
    msg.add_alternative(HTML_FRAME.format(content=content_html), subtype='html')

    # Отправка
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        smtp.send_message(msg)

    print(f"Письмо отправлено на {recipient_email}")


