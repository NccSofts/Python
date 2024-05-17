import pandas as pd
import pyodbc
import sqlalchemy
import  numpy as np
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

print("EXEC PowerBiv2..Mapa_Venda_Estoque")
cursor.execute("EXEC PowerBiv2..Rodar_MapaFiscal")


linha_vazia = '<p>&nbsp;</p>'

query = """\
            SELECT 
            
                CONCAT('Conciliação Faturamento: Foram encontradas ', COUNT(Numero), ' Nfs com diferença somando um total de ', COALESCE(SUM(Dif),0), ' - Posição em ',FORMAT(GETDATE(),'dd/MM/yyyy')) Titulo
                
            FROM PowerBIv2..Conciliacao_Faturamento
            WHERE CONCAT(Ano,Mes) = FORMAT(GETDATE(), 'yyyyMM') AND Status = 'Verificar'
        """

query2 = """\
            SELECT 
            
                COALESCE(ABS(SUM(Dif)),0) Total
                
            FROM PowerBIv2..Conciliacao_Faturamento
            WHERE Status = 'Verificar'
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
        <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Segue relação de diferenças encontradas na conciliação do faturamento entre fiscal e contábil.</strong></h1>
        <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Att.,</strong></h1>
        <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Equipe de BI Mais Proxima<br />&nbsp;</strong></h1>

        """

        with open("C:\Python\Contabilidade\Conciliacao_Faturamento.html", "w") as file:
            file.write(texto)

        with open("C:\Python\Contabilidade\Conciliacao_Faturamento.html", "a") as file:
            file.write(linha_vazia)

        with open("C:\Python\Contabilidade\Conciliacao_Faturamento.html", "a") as file:
            file.write('RESUMO CONCILIAÇÃO FATURAMENTO')

        pivot_sql = """\
                            SELECT 
                                                                                
                                Numero,
                                ForCli,
                                CONCAT(Ano,Mes) AnoMes,
                                StatusNF,
                                CFOP CFOP,
                                ROUND(Total_Fiscal,2) ValorFiscal,
                                ROUND(Total_Cont, 2) ValorContabil,
                                ROUND(Dif, 2) Diferença,
                                StatusDif Status
                              
                            FROM PowerBIv2..Base_Conciliacao_Faturamento                                                        
                            --FROM PowerBIv2..Conciliacao_Faturamento
                            WHERE  CONCAT(Ano,Mes) = FORMAT(GETDATE(), 'yyyyMM')
                            ORDER BY AnoMes, Numero
                    """
        pd.options.display.float_format = '{:,.2f}'.format
        df = pd.read_sql_query(pivot_sql, engine)
        df_resumo = pd.pivot_table(df, values=['ValorFiscal', 'ValorContabil', 'Diferença'], aggfunc=np.sum, index=['StatusNF', 'CFOP'], margins=True, margins_name='Total')
        df_resumo.to_html('C:\Python\Contabilidade\Resumo_Conciliacao_Faturamento.html')


        lista_sql = """\
        
                            SELECT 
                                                
                                    Numero,
                                    ForCli,
                                    CONCAT(Ano,Mes) AnoMes,
                                    StatusNF StatusNF,
                                    CFOP CFOP,ROUND(Dif, 2) Diferença,
                                    ROUND(Total_Cont, 2) ValorContabil,
                                    ROUND(Total_Fiscal ,2) ValorFiscal,
                                    StatusDif
                                                                
                            FROM PowerBIv2..Base_Conciliacao_Faturamento
                            WHERE StatusDif LIKE 'Verificar%' AND CONCAT(Ano,Mes) = FORMAT(GETDATE(), 'yyyyMM')
                            --ORDER BY AnoMes, Numero
                                                                
                            UNION ALL
                                                                
                            SELECT 
                                                                
                                '',
                                '',
                                '',
                                '',
                                'TOTAL GERAL',
                                COALESCE(ROUND(SUM(Total_Fiscal),2),0) ValorFiscal,
                                COALESCE(ROUND(SUM(Total_Cont), 2),0) ValorContabil,
                                COALESCE(ROUND(SUM(Dif), 2),0) Diferença,
                                ''
                                                                
                            FROM PowerBIv2..Base_Conciliacao_Faturamento
                            WHERE StatusDif LIKE 'Verificar%' AND CONCAT(Ano,Mes) = FORMAT(GETDATE(), 'yyyyMM')
                            --ORDER BY AnoMes, Numero ASC
                
                    """
        pd.options.display.float_format = '{:,.2f}'.format
        df2 = pd.read_sql_query(lista_sql, engine)
        df2.to_html('C:\Python\Contabilidade\Lista_Conciliacao_Faturamento.html')

        # Concatenando códigos HTML num unico arquivo
        tabela_sql = open('C:\Python\Contabilidade\Resumo_Conciliacao_Faturamento.html', 'r')
        html = tabela_sql.read()
        #
        with open("C:\Python\Contabilidade\Conciliacao_Faturamento.html", "a") as file:
            file.write(html)


        with open("C:\Python\Contabilidade\Conciliacao_Faturamento.html", "a") as file:
            file.write(linha_vazia)
        with open("C:\Python\Contabilidade\Conciliacao_Faturamento.html", "a") as file:
            file.write(linha_vazia)

        with open("C:\Python\Contabilidade\Conciliacao_Faturamento.html", "a") as file:
            file.write('LISTA DE NFs COM DIFERENÇA')

        # Concatenando códigos HTML num unico arquivo
        tabela_sql = open('C:\Python\Contabilidade\Lista_Conciliacao_Faturamento.html', 'r')
        html = tabela_sql.read()

        with open("C:\Python\Contabilidade\Conciliacao_Faturamento.html", "a") as file:
            file.write(html)




#
        corpo = codecs.open("C:\Python\Contabilidade\Conciliacao_Faturamento.html", 'r')
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
        # receiver_email = ['sheila.gomes@maisproxima.com.br', 'igor.chaves@maisproxima.com.br', 'contabilidade@maisproxima.com.br','Leonardo Oliveira <leonardo.oliveira@maisproxima.com.br>', 'ananda.silva@b2finance.com', 'Davi Leite Lemos <davi.lemos@b2finance.com>']
        receiver_email = ['logistica@maisproxima.com.br', 'igor.chaves@maisproxima.com.br',
                          'mailto:claudio.junior@b2finance.com',
                          'Andreia Santos <andreia.santos@maisproxima.com.br>',
                          'Davi Leite Lemos <davi.lemos@b2finance.com>', 'fiscal@maisproxima.com.br']
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
