import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
from os import getenv
from dotenv import load_dotenv

load_dotenv()


def send_email(subject, attachment):
    # Configurações do servidor SMTP do Gmail
    smtp_server = getenv('SMTP_SERVER')
    smtp_port = getenv('SMTP_PORT')
    smtp_username = getenv('EMAIL_ADDRESS')
    smtp_password = getenv('EMAIL_PASSWORD')

    # Criar objeto SMTP
    email_msg = MIMEMultipart()
    email_msg['From'] = smtp_username
    email_msg['To'] = getenv('EMAIL_TO')
    email_msg['Subject'] = subject
    body = f"""
    <html>
        <head>
            <style>
                p {{
                    font-size: 16px;
                    font-family: Arial, Helvetica, sans-serif;
                }}
            </style>
            
        </head>
        <body>
            <p>Olá,</p>
            <p>Segue em anexo a planilha com os dados dos processos.</p>
            <p>Att. {smtp_username}</p>
        </body>
    </html>
    """
    email_msg.attach(MIMEText(body, 'html'))

    # abrir arquivo
    if attachment:
        with open(attachment, 'rb') as file:
            att_mine = MIMEBase('application', 'octet-stream')
            att_mine.set_payload(file.read())
            encoders.encode_base64(att_mine)

            # Extrair o nome do arquivo do caminho completo
            attachment_name = os.path.basename(attachment)
            # Adicionar cabeçalho
            att_mine.add_header('Content-Disposition', f'attachment; filename={attachment_name}')

    # Colocar anexo no corpo do email
    email_msg.attach(att_mine)

    try:
        server = smtplib.SMTP(smtp_server, int(smtp_port))
        server.ehlo()  # Identifique-se com o servidor
        server.starttls()  # Inicie a criptografia TLS
        server.login(smtp_username, smtp_password)
        server.sendmail(email_msg['From'], email_msg['To'], email_msg.as_string())
        server.quit()
        print('Email enviado com sucesso!')
    except Exception as e:
        print(f'Algo deu errado... {e}')


