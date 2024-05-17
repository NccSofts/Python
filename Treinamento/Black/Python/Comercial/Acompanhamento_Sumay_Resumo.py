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
        <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Abaixo temos um resumo da Operação Sumay.</strong></h1>
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

resumo = pd.read_sql_table('Acompanhamento_Sumay_Resumo', engine)

# df = resumo
#
# fig = go.Figure(data=[go.Table(
#     header=dict(values=list(df.columns),
#                 fill_color='paleturquoise',
#                 align='left'),
#     cells=dict(values=[df.AnoMes],
#                fill_color='lavender',
#                align='left'))
# ])
#
# fig.show()

resumo.to_html('C:\Python\Comercial\Resumo1.html')

resumo_sql = open('C:\Python\Comercial\Resumo1.html', 'r', errors='ignore')
html = resumo_sql.read()

with open("C:\Python\Comercial\Acompanhamento_DL.html", "a") as file:
    file.write(html)

with open("C:\Python\Comercial\Acompanhamento_DL.html", "a") as file:
    file.write(linha_vazia)

lista_dl_faturar = """\

                            SELECT 

                                FORMAT(Data, 'yyyy') Ano,
                                FORMAT(Data, 'MM') Mes,
                                FORMAT(Data, 'dd/MMM') Dia,
                                Pedido,
                                UPPER([Razao Social]) [Razao Social],
                                [Total Pedido] 


                            FROM [DataWareHouse].[dbo].Industria_Pedido_NFe
                            WHERE Industria = 'Sumay' AND [Status-iMPX] LIKE '41-%' AND Data >= FORMAT(DATEADD(MONTH, -4, GETDATE()), 'yyyy-MM-01')

                            ORDER BY Data ASC, Pedido ASC

            """

lista_dl_pd = pd.read_sql_query(lista_dl_faturar, engine)
df = lista_dl_pd
dic = df.to_dict()

resumo_dl = pd.pivot_table(lista_dl_pd, values=['Total Pedido'], aggfunc='sum',
                           index=['Ano', 'Mes', 'Dia', 'Pedido', 'Razao Social'], margins=True, margins_name='Total',
                           fill_value=0)
resumo2_dl = pd.pivot_table(lista_dl_pd, values=['Total Pedido'], aggfunc='sum', index=['Ano', 'Mes'], margins=True,
                            margins_name='Total', fill_value=0)

resumo_dl.to_html('C:\Python\Comercial\Resumo1.html')
resumo2_dl.to_html('C:\Python\Comercial\Resumo2.html')

resumo_sql = open('C:\Python\Comercial\Resumo2.html', 'r', errors='ignore')
html = resumo_sql.read()

with open("C:\Python\Comercial\Acompanhamento_DL.html", "a") as file:
    file.write('<span style="text-decoration: underline;">RESUMO PEDIDOS À FATURAR:</span>')

with open("C:\Python\Comercial\Acompanhamento_DL.html", "a") as file:
    file.write(html)

with open("C:\Python\Comercial\Acompanhamento_DL.html", "a") as file:
    file.write(linha_vazia)

resumo_sql = open('C:\Python\Comercial\Resumo1.html', 'r', errors='ignore')
html = resumo_sql.read()

with open("C:\Python\Comercial\Acompanhamento_DL.html", "a") as file:
    file.write('<span style="text-decoration: underline;">RELAÇÃO DE PEDIDOS À FATURAR:</span>')

with open("C:\Python\Comercial\Acompanhamento_DL.html", "a") as file:
    file.write(html)

with open("C:\Python\Comercial\Acompanhamento_DL.html", "a") as file:
    file.write(linha_vazia)

lista_dl_faturar = """\

                        SELECT
                        
                            FORMAT(DL.[Data Faturamento], 'yyyy') Ano,
                            FORMAT(DL.[Data Faturamento], 'dd/MMM') Mes,
                            FORMAT(DL.[Data Faturamento], 'dd/MM/yyyy') Data,
                            Pedido,
                            FORMAT(Data, 'dd/MM/yyyy') DataPedido,
                            [Razao Social],
                            [Total Pedido] [Total Pedido]
                        
                        
                        FROM [DataWareHouse].[dbo].[Industria_Pedido_NFe] DL
                        INNER JOIN
                        (
                        SELECT TOP 3
                        
                            [Data Faturamento]
                        
                        FROM [DataWareHouse].[dbo].[Industria_Pedido_NFe]
                        WHERE Industria = 'Sumay' AND [Status-iMPX] LIKE '43-%'
                        GROUP BY [Data Faturamento]
                        ORDER BY [Data Faturamento] Desc
                        ) D ON D.[Data Faturamento] = DL.[Data Faturamento]
                        
                        WHERE Industria = 'Sumay' AND [Status-iMPX] LIKE '43-%' --AND Data >= FORMAT(DATEADD(MONTH, -4, GETDATE()), 'yyyy-MM-01')
                        
                        ORDER BY DL.[Data Faturamento] Desc, Dl.Pedido Desc
            """

lista_dl_pd = pd.read_sql_query(lista_dl_faturar, engine)

df = lista_dl_pd

resumo_dl = pd.pivot_table(lista_dl_pd, values=['Total Pedido'], aggfunc='sum',
                           index=['Ano', 'Data', 'Pedido', 'DataPedido' ,'Razao Social'], margins=True, margins_name='Total',
                           fill_value=0)

resumo_dl.to_html('C:\Python\Comercial\Resumo3.html')

resumo_sql = open('C:\Python\Comercial\Resumo3.html', 'r', errors='ignore')
html = resumo_sql.read()

with open("C:\Python\Comercial\Acompanhamento_DL.html", "a", ) as file:
    file.write('<span style="text-decoration: underline;">RESUMO PEDIDOS FATURADOS NOS ÚLTIMOS 3 DIAS:</span>')

with open("C:\Python\Comercial\Acompanhamento_DL.html", "a") as file:
    file.write(html)

with open("C:\Python\Comercial\Acompanhamento_DL.html", "a") as file:
    file.write(linha_vazia)

lista_pedidos_pendentes = """\

                            SELECT 
                                Pedido,
                                [Razao Social],
                                Data DataPedido,
                                FORMAT(ROUND(Valor,0), 'N', 'PT-BR') Valor,
                                AnoMes,
                                [Previsao Entrega],
                                RTRIM(SUBSTRING(Vendedor, 3 , 50)) Vendedor
                                
                            
                            FROM PowerBiv2..Acompanhamento_Sumay_Base
                            WHERE Status1 = 'Faturado' AND Status2 = 'Pendente Entrega'
                            ORDER BY Valor DESC
                            
                    """

pedidos_pendentes = pd.read_sql_query(lista_pedidos_pendentes, engine)

df = pedidos_pendentes

df.to_html('C:\Python\Comercial\Resumo4.html')

resumo_sql = open('C:\Python\Comercial\Resumo4.html', 'r', errors='ignore')
html = resumo_sql.read()

with open("C:\Python\Comercial\Acompanhamento_DL.html", "a") as file:
    file.write('<span style="text-decoration: underline;">PEDIDOS PENDENTES DE ENTREGA:</span>')

with open("C:\Python\Comercial\Acompanhamento_DL.html", "a") as file:
    file.write(html)

with open("C:\Python\Comercial\Acompanhamento_DL.html", "a") as file:
    file.write(linha_vazia)


with open("C:\Python\Comercial\Acompanhamento_DL.html", "a") as file:
    file.write('<span style="text-decoration:;">Este email é gerado automáticamente:</span>')

corpo = codecs.open("C:\Python\Comercial\Acompanhamento_DL.html", 'r')
html = corpo.read()
html = html.replace('<tr style="text-align: right;">', '<tr style="text-align: center;">')
html = html.replace('<td>', '<td style="text-align: center;">')




if diasemana < 5:

    # ENVIA EMAIL
    diasemana = date.today().weekday()
    data_atual = date.today()
    data_pt_br = data_atual.strftime('%d/%m/%Y')
    titulo = str('Acompanhamento Sumay - Posição em: ' + str(data_pt_br))

    # titulo = 'Acompanhamento Processos DL - Posição em ' + str(data_pt_br)
    # sender_email = "credito@maisproxima.com.br"
    user_login = 'bi@maisproxima.com'
    # receiver_email = ['anderson.souza@maisproxima.com.br']
    receiver_email = ['sheila.gomes@maisproxima.com.br', 'igor.chaves@maisproxima.com.br', 'Marcos Brando <marcos.brando@logare.com.br>', 'Simone Dario <simone.dario@maisproxima.com.br>', 'anderson.souza@maisproxima.com.br','Tatiana Mello <tatiana.mello@maisproxima.com.br>']
    # receiver_email = ['logistica@maisproxima.com.br', 'igor.chaves@maisproxima.com.br', 'claudio.junior@b2finance.com' ,'Andreia Santos <andreia.santos@maisproxima.com.br>' ,'Davi Leite Lemos <davi.lemos@b2finance.com>' ,'fiscal@maisproxima.com.br']
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

else:

    print("Emails enviados somente durante a semana")



