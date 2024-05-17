import os
import sys
import pandas as pd
import pyodbc
import sqlalchemy
import codecs
import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
import datetime
from datetime import timedelta, date, time
sys.path.append('C:/Python/Database/')
import database as db

diasemana = datetime.datetime.today().weekday()
hora_atual = str(datetime.datetime.today().time())

# Definir conexao 0 = PowerBiV2 / 1 = Datawarehouse
cursor = db.sql_conn("PowerBiv2")
engine = db.sql_engine("powerBiv2")
status_check_cursor = db.sql_conn("PowerBiv2")


print('Executando Onedrive_Faturas_Frete')
cursor.execute("EXEC PowerBiv2..Onedrive_Faturas_Frete")
while 1:
    q = status_check_cursor.execute("select RunningStatus from PowerBIv2..Acompanhamento_Rotinas_SQL_Python where Script_Procedure = 'Tabela_Onedrive_Faturas_Frete'").fetchone()
    if q[0] == 0:
        break

if diasemana < 5 and hora_atual <= '19:00':

    linha_vazia = '<p>&nbsp;</p>'

    texto = """\

    <!-- #######  YAY, I AM THE SOURCE EDITOR! #########-->
    <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Olá,</strong></h1>
    <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Segue relação de faturas da logistica para lançamento no Protheus.</strong></h1>
    <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Atentar para as faturas que não foram localizadas no Protheus.</strong></h1>
    <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Att.,</strong></h1>
    <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Equipe de BI Mais Proxima<br />&nbsp;</strong></h1>

    """

    with open("C:\Python\Fiscal\Faturas_Logistica.html", "w") as file:
        file.write(texto)

    with open("C:\Python\API_Facility\Faturas_Logistica.html", "a") as file:
        file.write(linha_vazia)

    df = pd.read_sql_table('vFaturas_Frete', engine)
    df = df.sort_values(by=['StatusFatura','Vencimento_Fatura'], ascending=False)
    # df = df.drop(columns="Order")
    df = df.reset_index()
    df = df.drop(columns="index")

    df.to_html('C:\Python\Fiscal\Lista_Faturas_Logistica.html')

    with open("C:\Python\Fiscal\Lista_Faturas_Logistica.html", "a") as file:
        file.write(linha_vazia)

    # Concatenando códigos HTML num unico arquivo
    tabela_sql = open('C:\Python\Fiscal\Lista_Faturas_Logistica.html', 'r')
    html = tabela_sql.read()

    with open("C:\Python\API_Facility\Faturas_Logistica.html", "a") as file:
        file.write(linha_vazia)

    with open("C:\Python\Fiscal\Faturas_Logistica.html", "a") as file:
        file.write(html)


    corpo = codecs.open("C:\Python\Fiscal\Faturas_Logistica.html", 'r')
    html = corpo.read()
    html = html.replace('<tr style="text-align: right;">', '<tr style="text-align: center;">')
    html = html.replace('<td>', '<td style="text-align: center;">')

    query_cabecalho = """\

                        SELECT 
                            CONCAT('Lista de Faturas Logsitca em ',FORMAT(GETDATE(), 'dd/MM/yyyy'),' - ',COUNT(Fatura), ' não encontrada(s) totalizando ', FORMAT(SUM(CAST(Valor_Frete AS float)), 'C', 'PT-BR')) Texto                  
                          FROM PowerBIv2..Tabela_Onedrive_FaturasFrete
                          WHERE DataProtheus IS NULL
                """
    cabecalho = pd.read_sql_query(query_cabecalho, engine)

    total = """\
                SELECT 
                    ABS(COALESCE(SUM([Valor_Frete]), 0)) Total
                FROM PowerBIv2..vFaturas_Frete
                """

    df = pd.read_sql_query(total, engine)
    Total = df['Total'].values[0]

    if Total > 0:

        # ENVIA EMAIL
        diasemana = date.today().weekday()
        data_atual = date.today()
        data_pt_br = data_atual.strftime('%d/%m/%Y')
        # sender_email = "credito@maisproxima.com.br"
        user_login = 'bi@maisproxima.com'
        # receiver_email = ['anderson.souza@maisproxima.com.br']
        receiver_email = ['anderson.souza@maisproxima.com.br', 'fiscal@maisproxima.com.br', 'adriano.goes@maisproxima.com.br','igor.chaves@maisproxima.com.br', 'Leonardo Oliveira <leonardo.oliveira@maisproxima.com.br>', 'Sheila Gomes <sheila.gomes@maisproxima.com.br>']
        # receiver_email = ['adriano.goes@maisproxima.com.br','anderson.souza@maisproxima.com.br']
        # receiver_email = ['logistica@maisproxima.com.br', 'igor.chaves@maisproxima.com.br','simone.dario@maisproxima.com.br', 'carlos.souza@maisproxima.com.br', 'tatiana.mello@maisproxima.com.br']
        # receiver_email = ['fiscal@maisproxima.com.br', 'igor.chaves@maisproxima.com.br']
        # password = input("Type your password and press enter:")
        # password = 'MP#qwert1234'
        password = '+Proxima2019'

        titulo_email = str(cabecalho['Texto'].values)
        titulo_email = titulo_email.replace("['",'')
        titulo_email = titulo_email.replace("']", '')

        sender_email = 'Gestão da Informação - Mais Próxima <bi@maisproxima.com>'
        # message = MIMEMultipart("alternative")
        message = MIMEMultipart("Related")
        message["Subject"] = titulo_email
        message["From"] = sender_email
        message["To"] = ", ".join(receiver_email)

        # Turn these into plain/html MIMEText objects
        # part1 = MIMEText(text, "plain")
        part2 = MIMEText(html, "html")

        # Add HTML/plain-text parts to MIMEMultipart message
        # The email client will try to render the last part first
        # message.attach(part1)
        message.attach(part2)

        # Imagem Contatos
        # This example assumes the image is in the current directory
        # fp = open('C:/Python/Gauge/Contatos.png', 'rb')
        # msgImage = MIMEImage(fp.read())
        # fp.close()

        # Define the image's ID as referenced above
        # msgImage.add_header('Content-ID', '<image1>')
        # message.attach(msgImage)

        # Imagem Meta
        # fp = open('C:/Python/Gauge/Meta.png', 'rb')
        # msgImage = MIMEImage(fp.read())
        # fp.close()
        #
        # msgImage.add_header('Content-ID', '<image2>')
        # message.attach(msgImage)

        # mailserver = smtplib.SMTP('smtp.office365.com', 587)
        mailserver = smtplib.SMTP('smtp.gmail.com', 587)
        mailserver.ehlo()
        mailserver.starttls()
        mailserver.login(user_login, password)
        mailserver.sendmail(sender_email, receiver_email, message.as_string())
        mailserver.quit()

        print('Email enviado com sucesso...')
    else:
        print('Não existem dados para envio do email')
else:
    print('Email enviado somente em dias uteis...')

