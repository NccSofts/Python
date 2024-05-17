import pandas as pd
import sqlalchemy
import pyodbc
import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import timedelta, date

# from bs4 import BeautifulSoup as Soup


# Definir conexao 0 = PowerBiV2 / 1 = Datawarehouse
opcao = 1


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


df = pd.read_sql_table('vBaseSimuleFreteErro_Frete', engine)
df = df.sort_values(by=['PedidoIMPX'])
df.to_html('C:\Python\API_Facility\Pedidos_Erro_Frete.html')

if len(df) > 0:

    print('Importando dados do SQL')

    data_atual = date.today()
    data_pt_br = data_atual.strftime('%d/%m/%Y')
    sender_email = 'edi.logistica@maisproxima.com.br'
    # receiver_email = "anderson.souza@bioscan.com.br"
    receiver_email = ['logistica@maisproxima.com.br', 'adriano.goes@maisproxima.com.br', 'fiscal@maisproxima.com.br',
                      'talita.amaral@maisproxima.com.br', 'carlos.souza@maisproxima.com.br', 'tatiana.mello@maisproxima.com.br']
    # receiver_email = ['anderson.souza@bioscan.com.br']
    # password = input("Type your password and press enter:")
    password = 'CFT654rfv'

    titulo_email = "Pedidos IMPX Cliente Retira que estão como CIF " + data_pt_br

    message = MIMEMultipart("alternative")
    message["Subject"] = titulo_email
    message["From"] = sender_email
    message["To"] = ", ".join(receiver_email)
    # message["To"] = receiver_email

    # Create the plain-text and HTML version of your message
    text = """\

    """

    texto = """\

    <!-- #######  YAY, I AM THE SOURCE EDITOR! #########-->
    <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Olá,</strong></h1>
    <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Favor verificar se os pedidos abaixo realmente devem ser CIF, pois usualmente são RETIRA/ES.</strong></h1>
    <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Att.,</strong></h1>
    <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Equipe de logistica Mais Proxima<br />&nbsp;</strong></h1>

    """

    # Concatenando códigos HTML num unico arquivo
    tabela_sql = open('C:\Python\API_Facility\Pedidos_Erro_Frete.html', 'r')
    html = tabela_sql.read()
    # encoding="utf-8"

    with open("C:\Python\API_Facility\corpoemail.html", "w") as file:
        file.write(texto)

    with open("C:\Python\API_Facility\corpoemail.html", "a") as file:
        file.write(html)

    # Recriando variaveis para envio do email
    tabela_sql = open('C:\Python\API_Facility\corpoemail.html', 'r')
    html = tabela_sql.read()

    print('Enviando Email')

    # Turn these into plain/html MIMEText objects
    part1 = MIMEText(text, "plain")
    part2 = MIMEText(html, "html")

    # Add HTML/plain-text parts to MIMEMultipart message
    # The email client will try to render the last part first
    message.attach(part1)
    message.attach(part2)

    # Create secure connection with server and send email
    # context = ssl.create_default_context()
    # with smtplib.SMTP_SSL("SMTP.office365.com", 587, context=context) as server:
    #     server.login(sender_email, password)
    #     server.sendmail(
    #         sender_email, receiver_email, message.as_string()
    #     )
    #
    mailserver = smtplib.SMTP('smtp.office365.com', 587)
    mailserver.ehlo()
    mailserver.starttls()
    mailserver.login(sender_email, password)
    mailserver.sendmail(sender_email, receiver_email, message.as_string())
    mailserver.quit()

    engine.execute("DELETE FROM DataWareHouse.dbo.Tabela_Simula_Frete WHERE NomeTranspSimulado LIKE '%errorMessage%'")
    print('')
    print('-----------------------------------------------------')
    print('-             EMAIL ENVIADO COM SUCESSO             -')
    print('-----------------------------------------------------')

    print('EXCLUINDO CLIENTES RETIRA DA BASE')
    cursor.execute("DELETE FROM DataWareHouse.dbo.Tabela_Simula_Frete WHERE ClienteRetira = 'RETIRA CD'")
    cursor.commit()

else:
    print('')
    print('-----------------------------------------------------')
    print('- EMAIL NÃO ENVIADO - NÃO EXISTEM REGISTRO COM ERRO -')
    print('-----------------------------------------------------')

