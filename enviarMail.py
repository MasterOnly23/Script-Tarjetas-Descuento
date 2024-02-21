import smtplib
import os
from decouple import config
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.text import MIMEText

def enviar_mail(file_path, date_str):
    try:
        fromaddr = config('ADDRESS_MAIL')
        password = config('PASSWORD_MAIL')
        toaddr = config('ADDRESS_MAIL_DEST').split(', ')

        msg = MIMEMultipart()

        msg['From'] = fromaddr
        msg['To'] = ", ".join(toaddr)
        msg['Subject'] = config('SUBJECT_MAIL') + date_str

        body = config('BODY_MAIL')

        msg.attach(MIMEText(body, 'plain'))

        filename = os.path.basename(file_path)
        attachment = open(file_path, "rb")

        part = MIMEBase('application', 'octet-stream')
        part.set_payload((attachment).read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

        msg.attach(part)

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(fromaddr, password)
        text = msg.as_string()
        server.sendmail(fromaddr, toaddr, text)
        server.quit()

        print("Mail enviado correctamente")

    except Exception as e:
        print("Ocurrió un error al enviar el correo electrónico: ", str(e))