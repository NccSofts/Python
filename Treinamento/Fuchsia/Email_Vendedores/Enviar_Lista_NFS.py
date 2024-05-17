import pandas as pd
import pyodbc
import sqlalchemy
from pandas import ExcelWriter
from email import encoders
import codecs
import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
import datetime
from datetime import timedelta, date

diasemana = datetime.datetime.today().weekday()



# Definir conexao 0 = PowerBiV2 / 1 = Datawarehouse
opcao = 0

configuracoes = pd.read_csv("C:/Python/database.cfg", header = 0, delimiter=";")

server = configuracoes['IP'].values[opcao]
database = configuracoes['Banco'].values[opcao]
username = configuracoes['Usuario'].values[opcao]
password = configuracoes['Senha'].values[opcao]


sql_string = 'mssql+pyodbc://' + username + ':' + password + '@' + server + '/' + database + '?driver=ODBC+Driver+13+for+SQL+Server'

cnxn = pyodbc.connect(
    'DRIVER={ODBC Driver 13 for SQL Server};SERVER=' + server + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password)
cursor = cnxn.cursor()

# Conexao com o servidor SQL
engine = sqlalchemy.create_engine(sql_string)
engine.connect()


df = pd.read_sql_table('Lista_NotasFiscais', engine)

if diasemana < 5:
    # GRAVA PLANILHA
    writer = ExcelWriter('C:\Python\API_Facility\Lista_Notas_Fiscais-MP.xlsx')
    df.to_excel(writer, 'Sheet1')
    writer.save()

    html = """\
            <p>Seguem em anexo rela&ccedil;&atilde;o de notas fiscais emitidas dentro do corrente m&ecirc;s.</p>
            <p>Att.,</p>
            <p>&nbsp;</p>
            <p><strong>Gest&atilde;o da Informa&ccedil;&atilde;o - Mais Pr&oacute;xima Distrinuidora.</strong></p>
            <p>&nbsp;</p>
            <p>&nbsp;</p>
        """


    # # ENVIA EMAIL
    diasemana = date.today().weekday()
    data_atual = date.today()
    data_pt_br = data_atual.strftime('%d/%m/%Y')
    # sender_email = "credito@maisproxima.com.br"
    user_login = 'bi@maisproxima.com'
    # receiver_email = ['anderson.souza@maisproxima.com.br']
    receiver_email = ['gustavo.vitali@facilitylog.net','Adriano Goes <adriano.goes@maisproxima.com.br>','Isaias Filho <isaias.filho@ifieldlogistics.com>','sergio.sanchez@facilitylog.net']
    # receiver_email = ['carlos.souza@maisproxima.com.br','simone.dario@maisproxima.com.br','tatiana.mello@maisproxima.com.br']
    # password = input("Type your password and press enter:")
    # password = 'MP#qwert1234'
    password = '+Proxima2019'

    titulo_email = 'LISTA DE NOTAS FISCAIS FATURADAS ATÉ ' + data_pt_br

    sender_email = 'Gestão da Informação - Mais Próxima <bi@maisproxima.com>'
    # message = MIMEMultipart("alternative")
    message = MIMEMultipart("Related")
    message["Subject"] = titulo_email
    message["From"] = sender_email
    message["To"] = ", ".join(receiver_email)
    #
    # # Turn these into plain/html MIMEText objects
    # # part1 = MIMEText(text, "plain")
    part2 = MIMEText(html, "html")
    message.attach(part2)

    part = MIMEBase('application', "octet-stream")
    part.set_payload(open('C:\Python\API_Facility\Lista_Notas_Fiscais-MP.xlsx', "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment; filename="Lista_Notas_Fiscais-MP.xlsx"')
    message.attach(part)
    #
    # # Add HTML/plain-text parts to MIMEMultipart message
    # # The email client will try to render the last part first
    # # message.attach(part1)
    # message.attach(part2)
    #
    # # Imagem Contatos
    # # This example assumes the image is in the current directory
    # # fp = open('C:/Python/Gauge/Contatos.png', 'rb')
    # # msgImage = MIMEImage(fp.read())
    # # fp.close()
    #
    # # Define the image's ID as referenced above
    # # msgImage.add_header('Content-ID', '<image1>')
    # # message.attach(msgImage)
    #
    # # Imagem Meta
    # # fp = open('C:/Python/Gauge/Meta.png', 'rb')
    # # msgImage = MIMEImage(fp.read())
    # # fp.close()
    # #
    # # msgImage.add_header('Content-ID', '<image2>')
    # # message.attach(msgImage)
    #
    # mailserver = smtplib.SMTP('smtp.office365.com', 587)
    mailserver = smtplib.SMTP('smtp.gmail.com', 587)
    mailserver.ehlo()
    mailserver.starttls()
    mailserver.login(user_login, password)
    mailserver.sendmail(sender_email, receiver_email, message.as_string())
    mailserver.quit()