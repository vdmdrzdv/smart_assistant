import smtplib
from email.message import EmailMessage

EMAIL_ADDRESS = "dlyashkolisusu@gmail.com"
EMAIL_PASSWORD = "rrkkuxqtzvecbwdb"

# Фиксированный HTML-шаблон с вставкой текста как блока
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

# Текстовая версия (резервная)
TEXT_TEMPLATE = """\
{content}
"""

def send_email_template(recipient_email, subject, body_message):
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = recipient_email

    msg.set_content(TEXT_TEMPLATE.format(content=body_message))
    msg.add_alternative(HTML_FRAME.format(content=body_message), subtype='html')

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        smtp.send_message(msg)

    print(f"Письмо отправлено на {recipient_email}")
