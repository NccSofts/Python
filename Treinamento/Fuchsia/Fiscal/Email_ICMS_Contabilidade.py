import os
import pandas as pd
import pyodbc
import sqlalchemy
import codecs
import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
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
    'DRIVER={ODBC Driver 13 for SQL Server};SERVER=' + server + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password, autocommit=True)
cursor = cnxn.cursor()

# Conexao com o servidor SQL
engine = sqlalchemy.create_engine(sql_string)
engine.connect()

if diasemana < 5:

    linha_vazia = '<p>&nbsp;</p>'

    query = """\
             SELECT 
            
                   [AnoMes]
                  ,FORMAT([Data], 'dd/MM/yyyy') Data
                  ,[Filial]
                  ,[Registro]
                  ,[NF_COMPRA]
                  ,[NF_REMESSA]
                  ,[DBCR]
                  ,[Valor]
                  ,Soma_Valor Diferença
                  ,[Historico]
                  ,Conciliado
                  ,Status
            
            FROM PowerBiv2..AnaliseICMS_Conciliado  
            WHERE Conciliado = 'Não' AND Historico NOT LIKE '%DEV NF 130935%' AND Historico NOT LIKE '%REMESSA NF 000362432%' 
									 AND Historico NOT LIKE '%DEV NF 131805%'
            ORDER BY NF_COMPRA
    """

    style_tabela = """\
    <style>
    table {
      border-collapse: collapse;
      width: 100%;
    }
    
    th, td {
      text-align: left;
      padding: 8px;
    }
    
    tr:nth-child(even) {background-color: #f2f2f2;}
    </style>
    """
    df = pd.read_sql_query(query, engine)
    df.to_html('C:\Python\Fiscal\Lista.html')

    #Concatenando códigos HTML num unico arquivo

    # Cabeçalho do Email
    # with open("C:\Python\Gauge\Resumo_Vendas.html", "w") as file:
    #     file.write(style_tabela)

    tabela_sql = open('C:\Python\Fiscal\Lista.html', 'r')
    html = tabela_sql.read()
    html = html.replace('<tr style="text-align: right;">','<tr style="text-align: center;">')
    html = html.replace('<td>','<td style="text-align: center;">')

    with open("C:\Python\Fiscal\Lista.html", "w") as file:
        file.write(html)

    corpo = codecs.open("C:\Python\Fiscal\Lista.html", 'r')
    html = corpo.read()
    html = html.replace('<tr style="text-align: right;">','<tr style="text-align: center;">')
    html = html.replace('<td>','<td style="text-align: center;">')

    # ENVIA EMAIL
    diasemana = date.today().weekday()
    data_atual = date.today()
    data_pt_br = data_atual.strftime('%d/%m/%Y')
    # sender_email = "credito@maisproxima.com.br"
    user_login = 'bi@maisproxima.com'
    # receiver_email = ['anderson.souza@maisproxima.com.br']
    # receiver_email = ['logistica@maisproxima.com.br']
    receiver_email = ['fiscal@maisproxima.com.br','igor.chaves@maisproxima.com.br']
    # password = input("Type your password and press enter:")
    # password = 'MP#qwert1234'
    password = '+Proxima2019'

    titulo_email = 'CONCILIAÇÃO ICMS CONTABILIDADE  - ' + data_pt_br

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
    print('Email enviado somente em dias uteis...')