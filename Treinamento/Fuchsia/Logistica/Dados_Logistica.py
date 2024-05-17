#!pip install gspread oauth2client


#IMPORTACAO PLANILHA FRETE COTACAO
import os
import gspread
import pyodbc
import sqlalchemy
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from pandas import DataFrame
from datetime import timedelta, date
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage



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



scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secrets.json', scope)
client = gspread.authorize(creds)

sheet1 = client.open('Expedicao_Recebimento').sheet1



Planilha1 = sheet1.get_all_records()

colunas1 = ['Data', 'Ano', 'Mes', 'Dia_Semana', 'Semana', 'Capacidade', 'QtdNF_Pend', 'QtdNF_Exp', 'Perc_Sep','CubEst']
df = pd.DataFrame.from_records(Planilha1, columns=colunas1)
df = df.astype(str)

print('IMPORTACAO PLANILHA DADOS OPERACAO LOGISTICA')
print('')
print('    Quantidade de registros a serem importados: '+ str(df.__len__()))
print('    Gravando na tabela SQL Tabela_Operacao_Logistica_Temp')

df.to_sql(
    name='Tabela_Operacao_Logistica_Temp', # database table name
    con=engine,
    if_exists='replace',
    index=False,
)

# # Grava na planilha que o registro foi importado para o SQL
# hoje = date.today()
# hojef = hoje.strftime('%d/%m/%Y')
# texto = 'Importado para SQL em '+str(hojef)
# lin = 2
# cell = 10
# for items in Planilha1:
#     sheet1.update_cell(lin, cell, texto)
#     lin = lin + 1

# IMPORTANDO STATUS DAS NFS PENDENTES DO DOCS
sheet2 = client.open('Expedicao_Recebimento').worksheet('2')
Planilha2 = sheet2.get_all_records()

colunas2 = ['NFs', 'OBS']
df = pd.DataFrame.from_records(Planilha2, columns=colunas2)
df = df.astype(str)

print('')
print('IMPORTACAO PLANILHA DADOS OPERACAO LOGISTICA STATUS NFS')
print('')
print('    Quantidade de registros a serem importados: '+ str(df.__len__()))
print('    Gravando na tabela SQL Tabela_Operacao_Logistica_Temp')

df.to_sql(
    name='Tabela_Operacao_Logistica2_Temp', # database table name
    con=engine,
    if_exists='replace',
    index=False,
)


# IMPORTANDO STATUS DAS NFS PENDENTES DO DOCS
sheet = client.open('Expedicao_Recebimento').worksheet('3')
Planilha = sheet.get_all_records()

colunas = ['Data', 'Ano', 'Mes', 'Dia_Semana', 'Semana', 'Capacidade', 'Carreta' ,'Truck' , 'Tres_Quartos', 'Total','QtdPltRec', 'QtdMP']
df = pd.DataFrame.from_records(Planilha, columns=colunas)
df = df.astype(str)

print('')
print('IMPORTACAO PLANILHA DADOS OPERACAO LOGISTICA STATUS NFS')
print('')
print('    Quantidade de registros a serem importados: '+ str(df.__len__()))
print('    Gravando na tabela SQL Tabela_Operacao_Logistica_Temp')

df.to_sql(
    name='Tabela_Operacao_Logistica3_Temp', # database table name
    con=engine,
    if_exists='replace',
    index=False,
)

# AJUSTANDO DADOS NAS TABELAS SQL
cursor.execute("EXEC OperacaoLog")
cursor.commit()


# CRIANDO GRAFICOS


# Variáveis para o Bar Chart NFS Expedidas
sql = pd.read_sql_query("SELECT LEFT(Data,2) Data, Capacidade*1 Capacidade, QtdNF_Exp*1 QtdNF_Exp FROM Tabela_Operacao_Logistica WHERE Ano = YEAR(GETDATE()) AND Mes = MONTH(GETDATE())", engine)

capacidade = sql.Capacidade
data_op = sql.Data
qtd_exp = sql.QtdNF_Exp
cont = 0


x = np.arange((len(data_op)))

bar_width = 0.40
fig = plt.figure(figsize=(10,4))
plt.bar(x, qtd_exp, width=bar_width, color='green', zorder=2)
sql['Capacidade'].plot(secondary_y=True, color='red')
plt.xticks(x + bar_width*2, data_op)
plt.title('Produtividade Operaçao Eximbiz')
plt.xlabel('Data')
plt.ylabel('Nf´s separadas no dia')
fig.savefig('C:\Python\Logistica\produtividade.png', dpi=100)
plt.close()
#plt.show()

#NFS Expedidas
df = pd.read_sql_table('OperacaoLog_PIVOT', engine)
df = df.fillna(0)
df.to_html('C:\Python\Logistica\OperacaoLog_PIVOT.html')



#NFs Pendentes de expediçao
df = pd.read_sql_table('OperacaoLog2_PIVOT', engine)
a = df.pivot_table(index=['Data'], columns=['OBS'], values=['VlrNF'], aggfunc={'VlrNF':[np.sum]}, fill_value=0)
a.to_html('C:\Python\Logistica\OperacaoLog2_PIVOT.html')



# Variáveis para o Bar Chart Carros Recebidos
query = "SELECT D.Ano,  D.Mes, LEFT(D.Data,2) as Dia, D.Capacidade, D.Total, T.QtdNF_Exp  FROM Tabela_Operacao_Logistica3 D\
        LEFT JOIN (SELECT Data, QTDNF_Exp FROM Tabela_Operacao_Logistica) T ON T.Data = D.Data\
        WHERE ANO = YEAR(GETDATE()) AND MES=MONTH(GETDATE())"
sql = pd.read_sql_query(query, engine)


capacidade = sql.Capacidade
dia_mes = sql.Dia
qtd_Rec = sql.Total
qtd_Exp = sql.QtdNF_Exp
cont = 0

x = np.arange((len(dia_mes)))

bar_width = 0.40
fig = plt.figure(figsize=(10,4))
#plt.bar(x, qtd_exp, width=bar_width, color='Blue', zorder=2)
sql['Total'].plot(kind='bar', color='blue')
sql['Capacidade'].plot(secondary_y=True, color='red')
plt.xticks(x + bar_width*2, dia_mes)
plt.title('Produtividade Operaçao Recebimento Eximbiz')
plt.xlabel('Dia/Mes')
plt.ylabel('Carros Recebidos')
plt.legend(loc='upper left')
#plt.gca().legend(('Carros Recebidos','Capacidade'))
fig.savefig('C:\Python\Logistica\produtividade2.png', dpi=100)
plt.close()


imagem1 = """\

<p><img src="https://ucf8e94d824c70c19d69f788d6b1.previews.dropboxusercontent.com/p/thumb/AAUUgWJxqvDd5_5-fU8f6pxlb7yK8Oie3itZZc6LziCRScmrYXuC0LTXhm6Ua1k7uV0O4N9wgwYug-WI_EkSCAWDFbuNZ0gefP4jf4zlsP18LXjnXYZpVJQs1m_WCMsVaucb9ecuLZQR5DL0bSEqyZtibM1RZw-sGC73XSfJNFLLHwMw1BnmNYadc0JNzF5UihPnmUy1Vu34K76GsJD1BezzQyydq2AHYBu7XqjT2CLG3en8loBvpGD_vBZaY9ebOFJ3o25LqHWhM7tu4l6EMjFU/p.png?size_mode=5" alt="" width="1000" height="400" /></p>

"""

imagem2 = """\

<p><img src="https://ucdde25e53a63f852425fef2876a.previews.dropboxusercontent.com/p/thumb/AAV-kSCtOcLBZDGqtQx0Vn19sUrzfpKoVmeC02FMofH2mptFF6p55c_czlzgdCHIzOF6PBdCWfn_2iHrjJBlqxBdjDXoJNjSGHh3m3hC9f_zLpQuaOrG6qp3RC2j68HeIdFqRZZWeYnQpvm2e0sCHiVcMAh9DQmmKGzSIY7FI4rgfKEorqJxDAG5XxW5uPoTfpC4bIU2HHWh5rAnkFPxnkvHfzC-igYEelIEY0hLTVtbjFvtYOIWUCU-ZKpGmS_i3CFgeX7c-aeGN_zT6dz77ecf/p.png?size=1600x1200&amp;size_mode=3" width="1000" height="400" /></p>

"""


texto = """\

<p><strong>Bom dia,</strong></p>
<p><strong>Segue abaixo resumo da opera&ccedil;&atilde;o logistica em Serra/ES.</strong></p>
<p><strong>Att.,</strong></p>
<p><strong>Equipe de Logistica - Mais Pr&oacute;xima</strong></p>
<p>&nbsp;</p>

"""

#Concatenando códigos HTML num unico arquivo
tabela_sql = open('C:\Python\Logistica\OperacaoLog_PIVOT.html', 'r')
html = tabela_sql.read()

with open("C:\Python\Logistica\corpoemail.html", "w") as file:
    file.write(texto)

with open("C:\Python\Logistica\corpoemail.html", "a") as file:
    file.write(imagem1)

with open("C:\Python\Logistica\corpoemail.html", "a") as file:
    file.write(html)

with open("C:\Python\Logistica\corpoemail.html", "a") as file:
    file.write(imagem2)

tabela_sql2 = open('C:\Python\Logistica\OperacaoLog2_PIVOT.html', 'r')
html2 = tabela_sql2.read()

with open("C:\Python\Logistica\corpoemail.html", "a") as file:
    file.write(html2)



#ENVIANDO EMAIL

# img_1 = open('C:\Python\Logistica\produtividade.png', 'rb').read()
# img_2 = open('C:\Python\Logistica\produtividade2.png', 'rb').read()
#
#
#
# #Recriando variaveis para envio do email
# tabela_sql = open('C:\Python\Logistica\corpoemail.html', 'r')
# html = tabela_sql.read()
# text = ''
#
# data_atual = date.today()
# data_pt_br = data_atual.strftime('%d/%m/%Y')
# sender_email = "edi.logistica@maisproxima.com.br"
# #receiver_email = "anderson.souza@bioscan.com.br"
# receiver_email = ['anderson.souza@mighty.com.br','igor.chaves@maisproxima.com.br','adriano.goes@maisproxima.com.br']
# #receiver_email = ['anderson.souza@bioscan.com.br']
# #password = input("Type your password and press enter:")
# password = 'Janeiro2019'
#
# titulo_email = "Posição Atualizada das Operações de Separação e Recebimento até " + data_pt_br
#
# message = MIMEMultipart("alternative")
# message["Subject"] = titulo_email
# message["From"] = sender_email
# message["To"] = ", ".join(receiver_email)
#
# # Turn these into plain/html MIMEText objects
# part1 = MIMEText(text, "plain")
# part2 = MIMEText(html, "html")
# image1 = MIMEImage(img_1, name='produtividade')
# image2 = MIMEImage(img_2, name='produtividade2')
#
# message.attach(part1)
# message.attach(part2)
# message.attach(image1)
# message.attach(image2)
#
# print('')
# print('')
# print('Enviando email')
#
# # Create secure connection with server and send email
# context = ssl.create_default_context()
# with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
#     server.login(sender_email, password)
#     server.sendmail(
#         sender_email, receiver_email, message.as_string()
#
#     )

print('')
print('')
print('Fim do Script')
