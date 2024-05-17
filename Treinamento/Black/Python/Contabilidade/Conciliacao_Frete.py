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

print("EXEC PowerBiV2..Proc_Razao_Contabil_Temp")
cursor.execute("EXEC PowerBiV2..Proc_Razao_Contabil_Temp")
while 1:
    q = status_check_cursor.execute("select RunningStatus from PowerBIv2..Acompanhamento_Rotinas_SQL_Python where Script_Procedure = 'Proc_Razao_Contabil_Temp'").fetchone()
    if q[0] == 0:
        break


print("EXEC PowerBIv2..Conciliacao_Frete")
comando = 'EXEC PowerBIv2..Conciliacao_Frete'
cursor.execute(comando)
# while 1:
#     q = status_check_cursor.execute("select RunningStatus from PowerBIv2..Acompanhamento_Rotinas_SQL_Python where Script_Procedure = 'Fechamento_Logistica'").fetchone()
#     if q[0] == 0:
#         break




linha_vazia = '<p>&nbsp;</p>'

query = """\
            SELECT 
            
            CONCAT('Conciliação Frete: Foram encontrados ', COUNT(CTE), ' CTE(s) com diferença somando um total de ', SUM(Dif), ' - Posição em ',FORMAT(GETDATE(),'dd/MM/yyyy')) Titulo
            
            FROM [PowerBIv2].[dbo].[Analise_Frete_Contabilidade]
            WHERE  Status_Dif <> 'OK' AND FORMAT(cast(IIF(Data_Cont IS NULL, Data_Fiscal, Data_Cont) as date), 'yyyyMM') = FORMAT(GETDATE(), 'yyyyMM')
        """


query2 = """\
            SELECT 
            
            COAlESCE(ABS(ROUND(SUM(Dif),0)),0) Total
            
            FROM [PowerBIv2].[dbo].[Analise_Frete_Contabilidade]
            WHERE  Status_Dif <> 'OK' AND FORMAT(cast(IIF(Data_Cont IS NULL, Data_Fiscal, Data_Cont) as date), 'yyyyMM') = FORMAT(GETDATE(), 'yyyyMM')
        """

# try:
df = pd.read_sql_query(query2, engine)
Total = df['Total'].values[0]

df = pd.read_sql_query(query, engine)
titulo = df['Titulo'].values[0]

if diasemana < 5:
    if Total > 0:

        texto = """\

        <!-- #######  YAY, I AM THE SOURCE EDITOR! #########-->
        <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Olá,</strong></h1>
        <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Segue relação de diferenças encontradas na conciliação das contas de frete entre fiscal e contábil.</strong></h1>
        <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Att.,</strong></h1>
        <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Equipe de BI Mais Proxima<br />&nbsp;</strong></h1>

        """

        with open("C:\Python\Contabilidade\Conciliacao_Frete.html", "w") as file:
            file.write(texto)

        with open("C:\Python\Contabilidade\Conciliacao_Frete.html", "a") as file:
            file.write(linha_vazia)

        lista_sql = """\
                    SELECT [CTE]
                            ,Data_Cont
                            ,Data_Fiscal
                            ,FORMAT(cast(IIF(Data_Cont IS NULL, Data_Fiscal, Data_Cont) as date), 'yyyyMM') AnoMes
                            ,COALESCE([Credito],0) Credito
                            ,COALESCE([Debito],0) Debito
                            ,COALESCE([Cancelado], 0) Cancelado
                            ,COALESCE([Total_Contabil], 0) Total_Contabil
                            ,[Total_Fiscal]
                            ,[Dif]
                            ,[Status_Dif]
                            ,[Analise]
                    FROM [PowerBIv2].[dbo].[Analise_Frete_Contabilidade]
                    WHERE  Status_Dif <> 'OK' AND FORMAT(cast(IIF(Data_Cont IS NULL, Data_Fiscal, Data_Cont) as date), 'yyyyMM') = FORMAT(GETDATE(), 'yyyyMM')
                    ORDER BY CTE ASC
        """


        df = pd.read_sql_query(lista_sql,engine)
        df.to_html('C:\Python\Contabilidade\Lista_Conciliacao_Frete.html')

        # Concatenando códigos HTML num unico arquivo
        tabela_sql = open('C:\Python\Contabilidade\Lista_Conciliacao_Frete.html', 'r')
        html = tabela_sql.read()

        with open("C:\Python\Contabilidade\Conciliacao_Frete.html", "a") as file:
            file.write(html)

        corpo = codecs.open("C:\Python\Contabilidade\Conciliacao_Frete.html", 'r')
        html = corpo.read()
        html = html.replace('<tr style="text-align: right;">', '<tr style="text-align: center;">')
        html = html.replace('<td>', '<td style="text-align: center;">')


        # ENVIA EMAIL
        diasemana = date.today().weekday()
        data_atual = date.today()
        data_pt_br = data_atual.strftime('%d/%m/%Y')
        # sender_email = "credito@maisproxima.com.br"
        user_login = 'bi@maisproxima.com'
        # receiver_email = ['anderson.souza@maisproxima.com.br']
        receiver_email = ['sheila.gomes@maisproxima.com.br', 'igor.chaves@maisproxima.com.br', 'contabilidade@maisproxima.com.br','Leonardo Oliveira <leonardo.oliveira@maisproxima.com.br>', 'Davi Leite Lemos <davi.lemos@b2finance.com>']
        # receiver_email = ['logistica@maisproxima.com.br', 'igor.chaves@maisproxima.com.br', 'claudio.junior@b2finance.com' ,'Andreia Santos <andreia.santos@maisproxima.com.br>' ,'Davi Leite Lemos <davi.lemos@b2finance.com>' ,'fiscal@maisproxima.com.br']
        # receiver_email = ['fiscal@maisproxima.com.br', 'igor.chaves@maisproxima.com.br']
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
        print("Não existem dados para serem enviados")
else:
    print("Emails enviados somente durante a semana")


# except:
#     print('Não existem dados para envio')
