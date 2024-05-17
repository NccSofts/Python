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




hoje = db.hoje()

diasemana = datetime.datetime.today().weekday()
hora_atual = str(datetime.datetime.today().time())

# Definir conexao 0 = PowerBiV2 / 1 = Datawarehouse
cursor = db.sql_conn("PowerBiv2")
engine = db.sql_engine("powerBiv2")
status_check_cursor = db.sql_conn("PowerBiv2")

global linha_vazia
global texto

if diasemana < 5:
    df1 = pd.read_sql('SELECT * FROM DataWareHouse.dbo.GrupoEconomico_A1(GETDATE())', engine)


    df2 = pd.read_sql('SELECT * FROM DataWareHouse.dbo.GrupoEconomico_A2', engine)
    df2.to_html("C:\Python\Comercial\GrupoEconomico_02.html", index=False)

    linha_vazia = '<p>&nbsp;</p>'

    texto = f"""\
            <p><img style="float: left;" src="cid:image1" alt="" width="250" height="65" /></p>
            <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Olá,</strong></h1>
            <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Segue análise de grupos econômicos em {hoje}.</strong></h1>
            <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Att.,</strong></h1>
            <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Equipe de BI Mais Proxima<br />&nbsp;</strong></h1>
    """

    with open("C:\Python\Comercial\GrupoEconomico.html", "w") as file:
        file.write(texto)

    with open("C:\Python\Comercial\GrupoEconomico.html", "a") as file:
        file.write(linha_vazia)

    texto = """\
                <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Análise Grupo Economico 01:</strong></h1>
            """
    with open("C:\Python\Comercial\GrupoEconomico.html", "a") as file:
        file.write(texto)

    if len(df1) > 0:

        df1.to_html("C:\Python\Comercial\GrupoEconomico_01.html", index=False)

        # Concatenando códigos HTML num unico arquivo
        tabela_sql = open("C:\Python\Comercial\GrupoEconomico_01.html", 'r')
        html = tabela_sql.read()
        html = html.replace('<tr style="text-align: right;">', '<tr style="text-align: center;">')
        html = html.replace('<td>', '<td style="text-align: center;">')

        with open("C:\Python\Comercial\GrupoEconomico.html", "a") as file:
            file.write(html)

        with open("C:\Python\Comercial\GrupoEconomico.html", "a") as file:
            file.write(linha_vazia)
    else:

        texto = """\
                    <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Não existem dados a serem mostrados.</strong></h1>
                """
        with open("C:\Python\Comercial\GrupoEconomico.html", "a") as file:
            file.write(texto)

        with open("C:\Python\Comercial\GrupoEconomico.html", "a") as file:
            file.write(linha_vazia)


    texto = """\
                <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Análise Grupo Economico 02:</strong></h1>
            """
    with open("C:\Python\Comercial\GrupoEconomico.html", "a") as file:
        file.write(texto)


    if len(df2) > 0:

        # Concatenando códigos HTML num unico arquivo
        tabela_sql = open("C:\Python\Comercial\GrupoEconomico_02.html", 'r')
        html = tabela_sql.read()
        html = html.replace('<tr style="text-align: right;">', '<tr style="text-align: center;">')
        html = html.replace('<td>', '<td style="text-align: center;">')
        html = html.replace('None', '')

        with open("C:\Python\Comercial\GrupoEconomico.html", "a") as file:
            file.write(html)

    else:

        texto = """\
                    <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Não existem dados a serem mostrados.</strong></h1>
                """
        with open("C:\Python\Comercial\GrupoEconomico.html", "a") as file:
            file.write(texto)


    # ENVIANDO EMAIL

    tabela_sql = open("C:\Python\Comercial\GrupoEconomico.html", 'r')
    html = tabela_sql.read()

    user_login = 'bi@maisproxima.com'
    password = '+Proxima2019'

    sender_email = 'Gestão da Informação - Mais Próxima <bi@maisproxima.com>'

    receiver_email = [
                        'ti@maisproxima.com.br',
                        'igor.chaves@maisproxima.com.br',
                        'simone.dario@maisproxima.com.br'
                     ]

    # receiver_email = ['logistica@maisproxima.com.br']

    titulo_email = 'Análise dos Grupos Econômicos em ' + str(hoje)

    message = MIMEMultipart("Related")
    message["Subject"] = titulo_email
    message["From"] = sender_email
    message["To"] = ", ".join(receiver_email)
    part2 = MIMEText(html, "html")
    message.attach(part2)

    # Imagem
    fp = open('C:/Python/Imagens/MPLogo.jpeg', 'rb')
    msgImage = MIMEImage(fp.read())
    fp.close()
    msgImage.add_header('Content-ID', '<image1>')
    message.attach(msgImage)

    try:
        mailserver = smtplib.SMTP('smtp.gmail.com', 587)
        mailserver.ehlo()
        mailserver.starttls()
        mailserver.login(user_login, password)
        mailserver.sendmail(sender_email, receiver_email, message.as_string())
        mailserver.quit()
        print('Email enviado com sucesso...')

    except Exception as error:
        error_string = str(error)
        nome_arquivo = str(os.path.basename(__file__))
        db.envia_email_erro(nome_arquivo, error_string)
        print('Email de erro enviado')

