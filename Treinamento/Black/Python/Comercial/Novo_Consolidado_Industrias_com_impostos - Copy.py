import ftplib
import pandas as pd
import pyodbc
import sqlalchemy
import numpy as np
import codecs
import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
import datetime
from datetime import timedelta, date, time

server = '52.67.196.1'
user = 'dl'
pswd = '7ubA/Wp2'

# server = '52.67.67.6'
# user = 'IMPX'
# pswd = 'P@ssw0rd'

# Open ftp connection
ftp = ftplib.FTP(server, user, pswd)

diasemana = datetime.datetime.today().weekday()
hora_atual = str(datetime.datetime.today().time())

# Definir conexao 0 = PowerBiV2 / 1 = Datawarehouse
opcao = 0

configuracoes = pd.read_csv("C:/Python/database.cfg", header=0, delimiter=";")

server = configuracoes['IP'].values[opcao]
database = configuracoes['Banco'].values[opcao]
username = configuracoes['Usuario'].values[opcao]
password = configuracoes['Senha'].values[opcao]

sql_string = 'mssql+pyodbc://' + username + ':' + password + '@' + server + '/' + database + '?driver=ODBC+Driver+13+for+SQL+Server'

cnxn = pyodbc.connect(
    'DRIVER={ODBC Driver 13 for SQL Server};SERVER=' + server + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password,
    autocommit=True)
cursor = cnxn.cursor()

status_check_cursor = pyodbc.connect(
    'DRIVER={ODBC Driver 13 for SQL Server};SERVER=' + server + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password,
    autocommit=True)
cursor2 = status_check_cursor.cursor()

# Conexao com o servidor SQL
engine = sqlalchemy.create_engine(sql_string)
engine.connect()

pd.options.display.float_format = '{:,.2f}'.format

linha_vazia = '<p>&nbsp;</p>'

texto = """\

        <!-- #######  YAY, I AM THE SOURCE EDITOR! #########-->
        <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Olá,</strong></h1>
        <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Abaixo temos um resumo das Operações das Indústrias.</strong></h1>
        <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Att.,</strong></h1>
        <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Equipe de BI Mais Proxima<br />&nbsp;</strong></h1>

        """

with open("C:\Python\Comercial\Acompanhamento_DL.html", "w") as file:
    file.write(texto)

with open("C:\Python\Comercial\Acompanhamento_DL.html", "a") as file:
    file.write(linha_vazia)

with open("C:\Python\Comercial\Acompanhamento_DL.html", "a") as file:
    file.write(linha_vazia)

with open("C:\Python\Comercial\Acompanhamento_DL.html", "a") as file:
    file.write('<span style="text-decoration: underline;">RESUMO:</span>')

resumo = pd.read_sql_table("Novo_Consolidado_Industrias", engine)
# resumo.drop('A Faturar', axis='columns', inplace=True)

resumo.to_html('C:\Python\Comercial\Resumo1.html', index=False)

resumo_sql = open('C:\Python\Comercial\Resumo1.html', 'r', errors='ignore')
html = resumo_sql.read()

with open("C:\Python\Comercial\Acompanhamento_DL.html", "a") as file:
    file.write(html)

with open("C:\Python\Comercial\Acompanhamento_DL.html", "a") as file:
    file.write(linha_vazia)

# MONTA TABELA PENDENTES DE ENTREGA EM ATRASO

pendentes_ent = """\
            SELECT * 
            FROM PowerBiv2..Acompanhamento_Industrias_Base_v2 
            WHERE AnoMes_Fat >= FORMAT(DATEADD(MONTH, -4, GETDATE()), 'yyyy-MM') AND
                    Status1 IN ('Pendente Entrega em atraso') AND 
                    Industria <> 'Positivo'	
        """
pendentes_entrega = pd.read_sql(pendentes_ent, engine)
resumo_pendentes = pd.pivot_table(pendentes_entrega, values=['Valor'], aggfunc='sum',
                                  index=['Industria', 'AnoMes_Fat', 'Pedido', 'Data_Pedido', 'Previsao Entrega', 'Razao_Social'], margins=True,
                                  margins_name='Total', fill_value=0)

resumo_pendentes.to_html('C:\Python\Comercial\Pendentes_Entrega.html')

with open("C:\Python\Comercial\Acompanhamento_DL.html", "a") as file:
    file.write('<span style="text-decoration: underline;">PEDIDOS PENDENTES DE ENTREGA EM ATRASO:</span>')

resumo_sql = open('C:\Python\Comercial\Pendentes_Entrega.html', 'r', errors='ignore')
html = resumo_sql.read()

with open("C:\Python\Comercial\Acompanhamento_DL.html", "a") as file:
    file.write(html)

with open("C:\Python\Comercial\Acompanhamento_DL.html", "a") as file:
    file.write(linha_vazia)

# PEDENTES DE FATURAMENTO
pendentes_fat = """\
            SELECT * 
            FROM PowerBiv2..Acompanhamento_Industrias_Base_v2 
            WHERE AnoMes_Fat >= FORMAT(DATEADD(MONTH, -4, GETDATE()), 'yyyy-MM') AND
                    Status1 IN ('A Faturar') AND 
                    StatusNaoAplicavel = 'Pendente Entrega'
        """
a_faturar = pd.read_sql(pendentes_fat, engine)
resumo_faturar = pd.pivot_table(a_faturar, values=['Valor'], aggfunc='sum',
                                  index=['Industria', 'AnoMes_Ped', 'Pedido', 'Data_Pedido', 'Razao_Social'], margins=True,
                                  margins_name='Total', fill_value=0)

resumo_faturar.to_html('C:\Python\Comercial\A_Faturar.html')

with open("C:\Python\Comercial\Acompanhamento_DL.html", "a") as file:
    file.write('<span style="text-decoration: underline;">PEDIDOS PENDENTES DE FATURAMENTO:</span>')

resumo_sql = open('C:\Python\Comercial\A_Faturar.html', 'r', errors='ignore')
html = resumo_sql.read()

with open("C:\Python\Comercial\Acompanhamento_DL.html", "a") as file:
    file.write(html)

with open("C:\Python\Comercial\Acompanhamento_DL.html", "a") as file:
    file.write(linha_vazia)



# PEDENTES DE ENVIO

pendentes_envio = """\
            SELECT * 
            FROM PowerBiv2..Acompanhamento_Industrias_Base_v2 
            WHERE AnoMes_Fat >= FORMAT(DATEADD(MONTH, -4, GETDATE()), 'yyyy-MM') AND 
            StatusNaoAplicavel IN ('Pendente de envio FIDC')
            ORDER BY Industria, Pedido
        """
envio_fidc = pd.read_sql(pendentes_envio, engine)
# resumo_fidc = pd.pivot_table(envio_fidc, values=['Valor'], aggfunc='sum',
#                                   index=['Industria', 'AnoMes_Ped', 'Pedido', 'Data_Pedido', 'Razao_Social'], margins=True,
#                                   margins_name='Total', fill_value=0)

resumo_fidc = pd.pivot_table(envio_fidc, values=['Valor'], aggfunc='sum',
                                  index=['Industria', 'AnoMes_Fat', 'NF', 'Data Faturamento', 'Pedido', 'Data Entrega', 'Razao_Social'], margins=True,
                                  margins_name='Total', fill_value=0)

resumo_fidc.to_html('C:\Python\Comercial\FIDC.html')

with open("C:\Python\Comercial\Acompanhamento_DL.html", "a") as file:
    file.write('<span style="text-decoration: underline;">PEDIDOS PENDENTES DE ENVIO FIDC:</span>')

resumo_sql = open('C:\Python\Comercial\FIDC.html', 'r', errors='ignore')
html = resumo_sql.read()

with open("C:\Python\Comercial\Acompanhamento_DL.html", "a") as file:
    file.write(html)

with open("C:\Python\Comercial\Acompanhamento_DL.html", "a") as file:
    file.write(linha_vazia)



corpo = codecs.open("C:\Python\Comercial\Acompanhamento_DL.html", 'r')
html = corpo.read()
html = html.replace('<tr style="text-align: right;">', '<tr style="text-align: center;">')
html = html.replace('<td>', '<td style="text-align: center;">')

if diasemana < 5:

    # ENVIA EMAIL
    diasemana = date.today().weekday()
    data_atual = date.today()
    data_pt_br = data_atual.strftime('%d/%m/%Y')
    titulo = str('Acompanhamento Consolidado de Serviços - Posição em: ' + str(data_pt_br))

    # titulo = 'Acompanhamento Processos DL - Posição em ' + str(data_pt_br)
    # sender_email = "credito@maisproxima.com.br"
    user_login = 'bi@maisproxima.com'
    # receiver_email = ['anderson.souza@maisproxima.com.br']
    receiver_email = ['sheila.gomes@maisproxima.com.br', 'igor.chaves@maisproxima.com.br', 'Simone Dario <simone.dario@maisproxima.com.br>',
                      'anderson.souza@maisproxima.com.br', 'Tatiana Mello <tatiana.mello@maisproxima.com.br>','Adriano Goes <adriano.goes@maisproxima.com.br>','Tesouraria Mais Próxima <tesouraria@maisproxima.com.br>']
    # # receiver_email = ['logistica@maisproxima.com.br', 'igor.chaves@maisproxima.com.br', 'claudio.junior@b2finance.com' ,'Andreia Santos <andreia.santos@maisproxima.com.br>' ,'Davi Leite Lemos <davi.lemos@b2finance.com>' ,'fiscal@maisproxima.com.br']
    # receiver_email = ['igor.chaves@maisproxima.com.br', 'Andreia Santos <andreia.santos@maisproxima.com.br>',
    #                   'sheila.gomes@maisproxima.com.br', 'anderson.souza@maisproxima.com.br',
    #                   'Artur Nucci Ferrari <ti@maisproxima.com.br>']
    # password = input("Type your password and press enter:")
    # password = 'MP#qwert1234'
    password = '+Proxima2019'

    # titulo_email = 'Análise Contas Transitórias em ' + str(data_pt_br)

    sender_email = 'Gestão da Informação - Mais Próxima <bi@maisproxima.com>'
    # message = MIMEMultipart("alternative")
    message = MIMEMultipart("Related")
    message["Subject"] = titulo
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

    print('Email Enviado com sucesso')

    cursor.execute("EXEC POWERBIV2..PROC_RUN_ROTINAS 'Novo Consolidado Industrias com impostos'")

else:

    print("Emails enviados somente durante a semana")



