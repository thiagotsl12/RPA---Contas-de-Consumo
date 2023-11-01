from datetime import datetime
from decouple import config
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

def send_email(subject, body, attachment_path):
    email_sup = config('EMAIL_SUP')
    email_dest = config('EMAIL_DEST')
    email_port = config('EMAIL_PORT')
    pass_sup = config('PASS_SUP')
    smtp_serv = config('SMTP_SERV')
    
    smtp_server = smtp_serv
    port = email_port
    enable_ssl = True
    username = email_sup
    password = pass_sup
    accept_untrusted_certificates = False
    send_from = email_sup
    send_to = email_dest
    
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = send_to
    msg['Subject'] = subject

    if body:
        msg.attach(MIMEText(body, 'plain'))
    
    # Adicionar a planilha como anexo
    if attachment_path:
        with open(attachment_path, "rb") as attachment:
            part = MIMEApplication(attachment.read(), _subtype="xlsx")
            part.add_header('content-disposition', 'attachment', filename=attachment_path)
            msg.attach(part)

    server = smtplib.SMTP(smtp_server, port)
    if enable_ssl:
        server.starttls()
    if not accept_untrusted_certificates:
        server.ehlo()
    server.login(username, password)
    server.sendmail(send_from, send_to, msg.as_string())
    server.quit()

if __name__ == "__main__":
    send_email()
