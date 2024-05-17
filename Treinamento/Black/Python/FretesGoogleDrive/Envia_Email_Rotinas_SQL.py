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
from datetime import timedelta, date, time

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
    'DRIVER={ODBC Driver 13 for SQL Server};SERVER=' + server + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password, autocommit=True)
cursor = cnxn.cursor()

status_check_cursor = pyodbc.connect(
    'DRIVER={ODBC Driver 13 for SQL Server};SERVER=' + server + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password, autocommit=True)
cursor2 = status_check_cursor.cursor()


# Conexao com o servidor SQL
engine = sqlalchemy.create_engine(sql_string)
engine.connect()


query = """\

        --CONFIRMA QUE ROTINA EXECUTOU 
        UPDATE PowerBiV2..Acompanhamento_Rotinas_SQL_Python

        SET Data = GETDATE(), Horario = FORMAT(GETDATE(), 'HH:mm'), Usuario = (SELECT CURRENT_USER), 
            [Status] = IIF((SELECT CURRENT_USER) = 'anderson.souza', 'Manual', 'Automatica'),
        Processo = (SELECT OBJECT_SCHEMA_NAME(@@PROCID) + '.' + OBJECT_NAME(@@PROCID)),
        Contador = Contador + 1
        WHERE Script_Procedure = 'Acompanhamento_Rotinas_SQL'


    """
cursor.execute(query)

query = """\

  UPDATE [PowerBIv2].[dbo].[Acompanhamento_Rotinas_SQL_Python]
		SET Status_Rotina = CASE
								WHEN Contador <> '0' AND Data = FORMAT(GETDATE(),'yyyy-MM-dd') THEN 'OK'
								ELSE 'VERIFICAR ROTINA'
							END
    """
cursor.execute(query)


if diasemana < 5 and hora_atual <= '19:00':

    linha_vazia = '<p>&nbsp;</p>'

    texto = """\

    <!-- #######  YAY, I AM THE SOURCE EDITOR! #########-->
    <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Olá,</strong></h1>
    <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Segue relação de rotinas que são automatizadas..</strong></h1>
    <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Atentar para as datas!!! .</strong></h1>
    <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Att.,</strong></h1>
    <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Equipe de BI Mais Proxima<br />&nbsp;</strong></h1>

    """

    with open("C:\Python\Fiscal\Rotinas_SQL.html", "w") as file:
        file.write(texto)

    with open("C:\Python\API_Facility\Rotinas_SQL.html", "a") as file:
        file.write(linha_vazia)

    df = pd.read_sql_table('Acompanhamento_Rotinas_SQL_Python', engine)
    df = df.sort_values(by=['Contador','Responsavel_Rotina'], ascending=True)
    df = df.drop(columns="RunningStatus")
    df = df.drop(columns="Descricao")
    df = df.drop(columns="Periodo_Rotina")
    df = df.reset_index()
    df = df.drop(columns="index")


    df.to_html('C:\Python\Fiscal\Lista_Rotinas_SQL.html')

    with open("C:\Python\Fiscal\Lista_Rotinas_SQL.html", "a") as file:
        file.write(linha_vazia)

    # Concatenando códigos HTML num unico arquivo
    tabela_sql = open('C:\Python\Fiscal\Lista_Rotinas_SQL.html', 'r')
    html = tabela_sql.read()

    with open("C:\Python\API_Facility\Rotinas_SQL.html", "a") as file:
        file.write(linha_vazia)

    with open("C:\Python\Fiscal\Rotinas_SQL.html", "a") as file:
        file.write(html)


    corpo = codecs.open("C:\Python\Fiscal\Rotinas_SQL.html", 'r')
    html = corpo.read()
    html = html.replace('<tr style="text-align: right;">', '<tr style="text-align: center;">')
    html = html.replace('<td>', '<td style="text-align: center;">')


    # ENVIA EMAIL
    diasemana = date.today().weekday()
    data_atual = date.today()
    data_pt_br = data_atual.strftime('%d/%m/%Y')
    # sender_email = "credito@maisproxima.com.br"
    # receiver_email = ['logistica@maisproxima.com.br']
    # receiver_email = ['anderson.souza@maisproxima.com.br', 'fiscal@maisproxima.com.br', 'adriano.goes@maisproxima.com.br','igor.chaves@maisproxima.com.br', 'Julio Sousa <julio.sousa@maisproxima.com.br>','Leonardo Oliveira <leonardo.oliveira@maisproxima.com.br>', 'Sheila Gomes <sheila.gomes@maisproxima.com.br>']
    receiver_email = ['anderson.souza@maisproxima.com.br','igor.chaves@maisproxima.com.br', 'ti@maisproxima.com.br']
    # receiver_email = ['logistica@maisproxima.com.br', 'igor.chaves@maisproxima.com.br','simone.dario@maisproxima.com.br', 'carlos.souza@maisproxima.com.br', 'tatiana.mello@maisproxima.com.br']
    # receiver_email = ['fiscal@maisproxima.com.br', 'igor.chaves@maisproxima.com.br']
    # password = input("Type your password and press enter:")
    # password = 'MP#qwert1234'
    user_login = 'bi@maisproxima.com'
    password = '+Proxima2019'

    query_titulo = """\

        SELECT 
	        IIF(COUNT(Status_Rotina) = 0, 'Todas as Rotinas foram executadas com sucesso',
            CONCAT('Acompanhamento de Rotinas Automatizadas SQL e Python em ',FORMAT(GETDATE(), 'dd/MM/yyyy'),' - ', COUNT(Status_Rotina), ' rotina(s) não foram executadas.')) Texto
    FROM [PowerBIv2].[dbo].[Acompanhamento_Rotinas_SQL_Python]
    WHERE Contador = 0
        """
    cabecalho = pd.read_sql_query(query_titulo, engine)




    # titulo_email = 'Acompanhamento de Rotinas Automatizadas SQL e Python em ' + str(data_pt_br)
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
    print('Email enviado somente em dias uteis...')



cursor.close()