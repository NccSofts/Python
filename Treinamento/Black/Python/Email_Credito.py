import pandas as pd
import sqlalchemy
import pyodbc
import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import timedelta, date
#from bs4 import BeautifulSoup as Soup


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


df = pd.read_sql_table('v_MapaCreditoSegurado2', engine)
df = df.sort_values(by=['ForCli'])
df.to_html('C:\Python\Tabela_segurados.html')
df2 = pd.read_sql_table('v_MapaCreditoSegurado4', engine)
df2.to_html('C:\Python\Tabela_segurados2.html')

print('Importando dados do SQL')

diasemana = date.today().weekday()
data_atual = date.today()
data_pt_br = data_atual.strftime('%d/%m/%Y')
sender_email = "credito@maisproxima.com.br"
# receiver_email = ['anderson.souza@maisproxima.com.br']
receiver_email = ['logistica@maisproxima.com.br','igor.chaves@maisproxima.com.br','sheila.gomes@maisproxima.com.br','simone.dario@maisproxima.com.br','carlos.souza@maisproxima.com.br','tatiana.mello@maisproxima.com.br','credito@maisproxima.com.br','cobranca@maisproxima.com.br','tesouraria@maisproxima.com.br']
#receiver_email = ['anderson.souza@bioscan.com.br']
#password = input("Type your password and press enter:")
password = 'MP#qwert1234'

titulo_email = "Posição dos Clientes com Crédito Segurado em " + data_pt_br

message = MIMEMultipart("alternative")
message["Subject"] = titulo_email
message["From"] = sender_email
message["To"] = ", ".join(receiver_email)
#message["To"] = receiver_email

# Create the plain-text and HTML version of your message
text = """\

"""

texto = """\

<!-- #######  YAY, I AM THE SOURCE EDITOR! #########-->
<h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Bom dia,</strong></h1>
<h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Segue relat&oacute;rio atualizado com a pos&ccedil;&atilde;o atual dos clientes com cr&eacute;dito segurado.</strong></h1>
<h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Att.,</strong></h1>
<h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Equipe de Cr&eacute;dito Mais Pr&oacute;xima<br />&nbsp;</strong></h1>

"""

#Concatenando códigos HTML num unico arquivo
tabela_sql = open('C:\Python\Tabela_segurados.html', 'r')
html = tabela_sql.read()
#encoding="utf-8"

with open("C:\Python\corpoemail.html", "w") as file:
    file.write(texto)

with open("C:\Python\corpoemail.html", "a") as file:
    file.write(html)


texto2 = """\

<p>&nbsp;</p>

"""

with open("C:\Python\corpoemail.html", "a") as file:
    file.write(texto2)

tabela_sql2 = open('C:\Python\Tabela_segurados2.html', 'r')
html2 = tabela_sql2.read()

with open("C:\Python\corpoemail.html", "a") as file:
    file.write(html2)


#Recriando variaveis para envio do email
tabela_sql = open('C:\Python\corpoemail.html', 'r')
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


if diasemana >= 0 < 5:
    mailserver = smtplib.SMTP('smtp.office365.com', 587)
    mailserver.ehlo()
    mailserver.starttls()
    mailserver.login(sender_email, password)
    mailserver.sendmail(sender_email, receiver_email, message.as_string())
    mailserver.quit()
else:
    print('Final de Semana - Email nao enviado')

