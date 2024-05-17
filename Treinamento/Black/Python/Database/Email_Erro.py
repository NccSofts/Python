import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import datetime
from datetime import timedelta, date, time

def hoje():
    return date.today().strftime('%d/%m/%Y')

today = hoje()

def envia_email_erro(script_python, msg_erro):

    global linha_vazia
    global texto

    user_login = 'bi@maisproxima.com'
    password = '+Proxima2019'
    sender_email = 'Gestão da Informação - Mais Próxima <bi@maisproxima.com>'
    receiver_email = ['anderson.souza@maisproxima.com.br']
    titulo_email = 'Erro ao rodar o script ' + str(script_python)

    html = f"""\

                <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Olá,</strong></h1>
                <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Houve um erro ao rodar o script {script_python} em {today}.</strong></h1>
                <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;"></strong></h1>
                <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;"></strong>Erro:</h1>
                <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;"></strong>{msg_erro}</h1>
    """

    message = MIMEMultipart("Related")
    message["Subject"] = titulo_email
    message["From"] = sender_email
    message["To"] = ", ".join(receiver_email)
    part = MIMEText(html, "html")
    message.attach(part)

    mailserver = smtplib.SMTP('smtp.gmail.com', 587)
    mailserver.ehlo()
    mailserver.starttls()
    mailserver.login(user_login, password)
    mailserver.sendmail(sender_email, receiver_email, message.as_string())
    mailserver.quit()

    return print('Email do erro enviado')
