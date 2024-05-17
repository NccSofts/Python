import ftplib
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

server = '52.67.196.1'
user = 'dl'
pswd = '7ubA/Wp2'


# server = '52.67.67.6'
# user = 'IMPX'
# pswd = 'P@ssw0rd'

#Open ftp connection
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
    'DRIVER={ODBC Driver 13 for SQL Server};SERVER=' + server + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password, autocommit=True)
cursor = cnxn.cursor()

status_check_cursor = pyodbc.connect(
    'DRIVER={ODBC Driver 13 for SQL Server};SERVER=' + server + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password, autocommit=True)
cursor2 = status_check_cursor.cursor()


# Conexao com o servidor SQL
engine = sqlalchemy.create_engine(sql_string)
engine.connect()


t_nao_gerados = """\

        SELECT 
               ZX3.ZX3_FILIAL Filial, 
               ZX3.ZX3_NF NF, 
               ZX3.ZX3_SERIE Serie, 
               ZX1.ZX1_PEDIDO Pedido, 
               ZX1.ZX1_COND CondPag, 
               SA1.A1_COD Codigo, 
               SA1.A1_LOJA Loja, 
               SA1.A1_NREDUZ Cliente, 
               ZX3.ZX3_PARC Parcela, 
               ZX1.ZX1_DATA Emissao, 
               ZX3.ZX3_VENCTO Vencto, 
               ZX3.ZX3_DIFDIA Dias, 
               ZX3.ZX3_DTENTR Entrega, 
               ZX3.ZX3_VALOR Valor, 
               ZX3.R_E_C_N_O_ NFEREGTO 
        FROM PJKTB2..ZX3080 ZX3 
            INNER JOIN PJKTB2..ZX1080 ZX1 ON 
                    ZX1.D_E_L_E_T_='' 
                    AND ZX1.ZX1_FILIAL=ZX3.ZX3_FILIAL 
                    AND ZX1.ZX1_NF=ZX3.ZX3_NF 
                    AND ZX1.ZX1_SERIE=ZX3.ZX3_SERIE 
                    AND ZX1.ZX1_PEDIDO=ZX3.ZX3_PEDIDO 
                    AND ZX1.ZX1_AVISTA='N'  
            INNER JOIN PJKTB2..SA1080 SA1 ON 
                    SA1.D_E_L_E_T_='' 
                    AND A1_CGC=ZX1.ZX1_CNPJ 
        WHERE ZX3.D_E_L_E_T_=''  
                    AND ZX3.ZX3_CR='T'  
                    AND ZX3.ZX3_DTENTR<>'' 
                    AND ZX3.ZX3_NF NOT IN 
                        (
                        SELECT E1_NUM
                        FROM PJKTB2..SE1080
                        WHERE SE1080.D_E_L_E_T_=''
                        )
"""

df_t_nao_gerados = pd.read_sql_query(t_nao_gerados, engine)

tt_nao_gerados = """\

                SELECT 
                
                       COALESCE(SUM(ZX3.ZX3_VALOR),0) Valor
                FROM PJKTB2..ZX3080 ZX3 
                    INNER JOIN PJKTB2..ZX1080 ZX1 ON 
                            ZX1.D_E_L_E_T_='' 
                            AND ZX1.ZX1_FILIAL=ZX3.ZX3_FILIAL 
                            AND ZX1.ZX1_NF=ZX3.ZX3_NF 
                            AND ZX1.ZX1_SERIE=ZX3.ZX3_SERIE 
                            AND ZX1.ZX1_PEDIDO=ZX3.ZX3_PEDIDO 
                            AND ZX1.ZX1_AVISTA='N'  
                    INNER JOIN PJKTB2..SA1080 SA1 ON 
                            SA1.D_E_L_E_T_='' 
                            AND A1_CGC=ZX1.ZX1_CNPJ 
                WHERE ZX3.D_E_L_E_T_=''  
                            AND ZX3.ZX3_CR='T'  
                            AND ZX3.ZX3_DTENTR<>'' 
                            AND ZX3.ZX3_NF NOT IN 
                                (
                                SELECT E1_NUM
                                FROM PJKTB2..SE1080
                                WHERE SE1080.D_E_L_E_T_=''
                                )
                """

df_tt_nao_gerados = pd.read_sql_query(tt_nao_gerados, engine)


t_pendentes = """\

                SELECT 
                    ZX3.ZX3_FILIAL Filial, 
                    ZX3.ZX3_NF NF, 
                    ZX3.ZX3_SERIE Serie, 
                    ZX1.ZX1_PEDIDO Pedido, 
                    ZX1.ZX1_COND CondPag, 
                    SA1.A1_COD Codigo, 
                    SA1.A1_LOJA Loja, 
                    SA1.A1_NREDUZ Cliente, 
                    ZX3.ZX3_PARC Parcela, 
                    ZX1.ZX1_DATA Emissao, 
                    ZX3.ZX3_VENCTO Vencto, 
                    ZX3.ZX3_DIFDIA Dias, 
                    ZX3.ZX3_DTENTR Entrega, 
                    ZX3.ZX3_VALOR Valor, 
                    ZX3.R_E_C_N_O_ NFEREGTO 
                FROM PJKTB2..ZX3080 ZX3 
                    INNER JOIN PJKTB2..ZX1080 ZX1 ON 
                            ZX1.D_E_L_E_T_='' 
                            AND ZX1.ZX1_FILIAL=ZX3.ZX3_FILIAL 
                            AND ZX1.ZX1_NF=ZX3.ZX3_NF 
                            AND ZX1.ZX1_SERIE=ZX3.ZX3_SERIE 
                            AND ZX1.ZX1_PEDIDO=ZX3.ZX3_PEDIDO 
                            AND ZX1.ZX1_VALIDA='F' 
                            AND ZX1.ZX1_PROC='F'  
                            AND ZX1.ZX1_AVISTA='N'  
                            AND ZX1.ZX1_DTENTR<>''  
                    INNER JOIN PJKTB2..SA1080 SA1 ON 
                            SA1.D_E_L_E_T_='' 
                            AND A1_CGC=ZX1.ZX1_CNPJ 
                WHERE ZX3.D_E_L_E_T_=''  
                            AND ZX3.ZX3_CR='F'  
                            AND ZX3.ZX3_DTENTR<>''
                """
df_t_pendentes = pd.read_sql_query(t_pendentes, engine)


tt_pendentes = """\

                SELECT 
                
                    COALESCE(SUM(ZX3.ZX3_VALOR),0) Valor
                FROM PJKTB2..ZX3080 ZX3 
                    INNER JOIN PJKTB2..ZX1080 ZX1 ON 
                            ZX1.D_E_L_E_T_='' 
                            AND ZX1.ZX1_FILIAL=ZX3.ZX3_FILIAL 
                            AND ZX1.ZX1_NF=ZX3.ZX3_NF 
                            AND ZX1.ZX1_SERIE=ZX3.ZX3_SERIE 
                            AND ZX1.ZX1_PEDIDO=ZX3.ZX3_PEDIDO 
                            AND ZX1.ZX1_VALIDA='F' 
                            AND ZX1.ZX1_PROC='F'  
                            AND ZX1.ZX1_AVISTA='N'  
                            AND ZX1.ZX1_DTENTR<>''  
                    INNER JOIN PJKTB2..SA1080 SA1 ON 
                            SA1.D_E_L_E_T_='' 
                            AND A1_CGC=ZX1.ZX1_CNPJ 
                WHERE ZX3.D_E_L_E_T_=''  
                            AND ZX3.ZX3_CR='F'  
                            AND ZX3.ZX3_DTENTR <>''
                """

df_tt_pendentes = pd.read_sql_query(tt_pendentes, engine)




linha_vazia = '<p>&nbsp;</p>'

texto = """\

        <!-- #######  YAY, I AM THE SOURCE EDITOR! #########-->
        <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Olá,</strong></h1>
        <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Abaixo temos uma analise dos processos entre DL e Mais Próxima.</strong></h1>
        <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Att.,</strong></h1>
        <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Equipe de BI Mais Proxima<br />&nbsp;</strong></h1>

        """

with open("C:\Python\Comercial\Arquivos_FTP_DL.html", "w") as file:
    file.write(texto)

with open("C:\Python\Comercial\Arquivos_FTP_DL.html", "a") as file:
    file.write(linha_vazia)

with open("C:\Python\Comercial\Arquivos_FTP_DL.html", "a") as file:
    file.write(linha_vazia)

with open("C:\Python\Comercial\Arquivos_FTP_DL.html", "a") as file:
    file.write('<span style="text-decoration: underline;">VERIFICAÇÃO NFs e TÍTULOS DL:</span>')
with open("C:\Python\Comercial\Arquivos_FTP_DL.html", "a") as file:
    file.write(linha_vazia)


Total = df_tt_pendentes['Valor'].values[0]

if Total > 0:

    print('NFs Pendentes para Geração de Títulos no Contas a Receber:')
    print(df_t_pendentes)

    with open("C:\Python\Comercial\Arquivos_FTP_DL.html", "a") as file:
        file.write('NFs Pendentes para Geração de Títulos no Contas a Receber:')

    df_t_pendentes.to_html('C:\Python\Comercial\Lista_Titulos_Pendentes.html')

    tabela_sql = open('C:\Python\Comercial\Lista_Titulos_Pendentes.html', 'r')
    html = tabela_sql.read()

    with open("C:\Python\Comercial\Arquivos_FTP_DL.html", "a") as file:
        file.write(html)
    with open("C:\Python\Comercial\Arquivos_FTP_DL.html", "a") as file:
        file.write(linha_vazia)

else:
    with open("C:\Python\Comercial\Arquivos_FTP_DL.html", "a") as file:
        file.write('Não existem NFs Pendentes para Geração de Títulos no Contas a Receber')

    with open("C:\Python\Comercial\Arquivos_FTP_DL.html", "a") as file:
        file.write(linha_vazia)




Total2 = df_tt_nao_gerados['Valor'].values[0]

if Total2 > 0:
    print('Lista de Nfs para emissão de títulos:')
    print(df_t_nao_gerados)

    with open("C:\Python\Comercial\Arquivos_FTP_DL.html", "a") as file:
        file.write('Lista de Nfs para emissão de títulos:')

    df_t_nao_gerados.to_html('C:\Python\Comercial\Lista_Titulos_NG.html')

    tabela_sql2 = open('C:\Python\Comercial\Lista_Titulos_NG.html', 'r')
    html2 = tabela_sql2.read()

    with open("C:\Python\Comercial\Arquivos_FTP_DL.html", "a") as file:
        file.write(html2)

    with open("C:\Python\Comercial\Arquivos_FTP_DL.html", "a") as file:
        file.write(linha_vazia)


else:
    with open("C:\Python\Comercial\Arquivos_FTP_DL.html", "a") as file:
        file.write(linha_vazia)
    with open("C:\Python\Comercial\Arquivos_FTP_DL.html", "a") as file:
        file.write('Não existem NFs pendentes para gerar titulos')
    with open("C:\Python\Comercial\Arquivos_FTP_DL.html", "a") as file:
        file.write(linha_vazia)




entrega = []
faturamento = []

ftp.cwd("/arquivos/upload/entrega")

try:
    entrega = ftp.nlst()
except ftplib.error_perm as resp:
    if str(resp) == "550 No files found":
        print("Não existem arquivos na pasta")
    else:
        raise


with open("C:\Python\Comercial\Arquivos_FTP_DL.html", "a") as file:
    file.write('<span style="text-decoration: underline;">VERIFICAÇÃO DOS ARQUIVOS DL NO FTP:</span>')
with open("C:\Python\Comercial\Arquivos_FTP_DL.html", "a") as file:
    file.write(linha_vazia)

if entrega[0] != 'old':

    with open("C:\Python\Comercial\Arquivos_FTP_DL.html", "a") as file:
        file.write('<strong>Lista de arquivos parados na pasta: Entrega</strong>')

    print("Lista de arquivos parados na pasta: Entrega")

    for f in entrega:


        arq = '<p> ' + f + '</p>'

        if f.find('old') == 0:
            print('pasta old ignorada...')
        else:
            with open("C:\Python\Comercial\Arquivos_FTP_DL.html", "a") as file:
                file.write(arq)
            print(f)

    with open("C:\Python\Comercial\Arquivos_FTP_DL.html", "a") as file:
        file.write(linha_vazia)


else:
    print("Nenhum arquivo na pasta Entrega a ser listado")
    with open("C:\Python\Comercial\Arquivos_FTP_DL.html", "a") as file:
        file.write('<strong>Nenhum arquivo na pasta Entrega a ser listado</strong>')

    with open("C:\Python\Comercial\Arquivos_FTP_DL.html", "a") as file:
        file.write(linha_vazia)



ftp.cwd("/arquivos/upload/faturamento")


try:
    faturamento = ftp.nlst()
except ftplib.error_perm as resp:
    if str(resp) == "550 No files found":
        print("Não existem arquivos na pasta")
    else:
        raise

if faturamento[0] != 'old':

    with open("C:\Python\Comercial\Arquivos_FTP_DL.html", "a") as file:
        file.write('<strong>Lista de arquivos parados na pasta: Faturamento</strong>')


    print("Lista de arquivos parados na pasta: Faturamento")

    for f in faturamento:

        arq = '<p> ' + f + '</p>'

        if f.find('old') == 0:
            print('pasta old ignorada...')
        else:
            with open("C:\Python\Comercial\Arquivos_FTP_DL.html", "a") as file:
                file.write(arq)
            print(f)
else:

    print("Nenhum arquivo na pasta Faturamento a ser listado")

    with open("C:\Python\Comercial\Arquivos_FTP_DL.html", "a") as file:
        file.write('<strong>Nenhum arquivo na pasta Faturamento a ser listado</strong>')


with open("C:\Python\Comercial\Arquivos_FTP_DL.html", "a") as file:
    file.write(linha_vazia)

with open("C:\Python\Comercial\Arquivos_FTP_DL.html", "a") as file:
    file.write('<span style="text-decoration: underline;">RESUMO NFs COM PRAZO DE ENTREGA EXPIRADO E SEM CONFIRMAÇÃO DE ENTREGA:</span>')
with open("C:\Python\Comercial\Arquivos_FTP_DL.html", "a") as file:
    file.write(linha_vazia)

entregadl_query = """\

                    SELECT
                                                                            
                        FORMAT([Data Faturamento], 'yyyy-MM') AnoMes,
                        COUNT(DISTINCT NF) Qtd_Nfs,
                        FORMAT(SUM([Total Pedido]), 'N', 'PT-BR') ValorTotal
                                                                            
                                                                            
                    FROM DataWareHouse..Industria_Pedido_NFe
                    WHERE Industria = 'DL' 
                            AND [Status-iMPX] = '43-Faturado'
                            AND [Status-Industria-NF] NOT IN ('80-EXTRAVIO\CANCELADA','50-CANCELADO - NF', '70-DEVOLVIDO')
                            --AND [Status-DL Pedido] NOT LIKE '60-ENTREGUE%'
                            AND [A Vista]= 'N'
                            AND [Data Entrega] IS NULL
                            AND [Previsao Entrega] <= DATEADD(DAY, -1, GETDATE())
                                                                            
                    GROUP BY FORMAT([Data Faturamento], 'yyyy-MM')
                                                                            
                    UNION ALL
                                                                            
                    SELECT
                                                                            
                        'Total Geral' AnoMes,
                        COUNT(DISTINCT NF) Qtd_Nfs,
                        FORMAT(SUM([Total Pedido]), 'N', 'PT-BR') ValorTotal
                                                                            
                                                                            
                    FROM DataWareHouse..Industria_Pedido_NFe
                    WHERE Industria = 'DL'
                            AND [Status-iMPX] = '43-Faturado'
                            --AND [Status-DL Pedido] NOT LIKE '60-ENTREGUE%'
                            AND [Status-Industria-NF] NOT IN ('80-EXTRAVIO\CANCELADA','50-CANCELADO - NF', '70-DEVOLVIDO')
                            AND [A Vista]= 'N'
                            AND [Data Entrega] IS NULL
                            AND [Previsao Entrega] <= DATEADD(DAY, -1, GETDATE())

        """

entregadl_resumo = pd.read_sql_query(entregadl_query, engine)

entregadl_resumo.to_html('C:\Python\Comercial\Resumo_EntregaDL.html')

# Concatenando códigos HTML num unico arquivo
tabela_sql = open('C:\Python\Comercial\Resumo_EntregaDL.html', 'r')
html = tabela_sql.read()

with open("C:\Python\Comercial\Arquivos_FTP_DL.html", "a") as file:
    file.write(html)

with open("C:\Python\Comercial\Arquivos_FTP_DL.html", "a") as file:
    file.write(linha_vazia)
with open("C:\Python\Comercial\Arquivos_FTP_DL.html", "a") as file:
    file.write('Lista de NFs com prazo de entrega expirado:')

entregadl_query = """\

                        SELECT
                                                                        
                            NF,
                            [Data Faturamento],
                            [Previsao Entrega],
                            [CNPJ-CPF],
                            [Razao Social],
                            UF,
                            FORMAT([Total Pedido], 'N', 'PT-BR') [Total Pedido]
                                                                                                
                        FROM DataWareHouse..Industria_Pedido_NFe
                        WHERE Industria = 'DL'
                                AND [Status-iMPX] = '43-Faturado'
                                --AND [Status-DL Pedido] NOT LIKE '60-ENTREGUE%'
                                AND [Status-Industria-NF] NOT IN ('80-EXTRAVIO\CANCELADA','50-CANCELADO - NF', '70-DEVOLVIDO')
                                AND [A Vista]= 'N'
                                AND [Data Entrega] IS NULL
                                AND [Previsao Entrega] <= DATEADD(DAY, -1, GETDATE())
                                                                                                
                        ORDER BY [Previsao Entrega] ASC

                """

entregadl_resumo = pd.read_sql_query(entregadl_query, engine)

entregadl_resumo.to_html('C:\Python\Comercial\Lista_EntregaDL.html')

# Concatenando códigos HTML num unico arquivo
tabela_sql = open('C:\Python\Comercial\Lista_EntregaDL.html', 'r')
html = tabela_sql.read()

with open("C:\Python\Comercial\Arquivos_FTP_DL.html", "a") as file:
    file.write(html)

corpo = codecs.open("C:\Python\Comercial\Arquivos_FTP_DL.html", 'r')
html = corpo.read()

query = """\

                SELECT
                                        
                    CONCAT('Acompanhamento Processos DL / Foram encontradas ', COUNT(NF), ' nfs que somam um total de ', FORMAT(SUM([Total Pedido]), 'C', 'PT-BR')) Titulo
                                                                                                     
                FROM DataWareHouse..Industria_Pedido_NFe
                WHERE	Industria = 'DL'
                        AND [Status-iMPX] = '43-Faturado'
                        --AND [Status-DL Pedido] NOT LIKE '60-ENTREGUE%'
                        AND [Status-Industria-NF] NOT IN ('80-EXTRAVIO\CANCELADA','50-CANCELADO - NF', '70-DEVOLVIDO')
                        AND [A Vista]= 'N'
                        AND [Data Entrega] IS NULL
                        AND [Previsao Entrega] <= DATEADD(DAY, -1, GETDATE())

                """

df = pd.read_sql_query(query, engine)
titulo = df['Titulo'].values[0]


if diasemana < 5:

    # ENVIA EMAIL
    diasemana = date.today().weekday()
    data_atual = date.today()
    data_pt_br = data_atual.strftime('%d/%m/%Y')
    titulo = str(titulo)
    # titulo = 'Acompanhamento Processos DL - Posição em ' + str(data_pt_br)
    # sender_email = "credito@maisproxima.com.br"
    user_login = 'bi@maisproxima.com'
    # receiver_email = ['anderson.souza@maisproxima.com.br']
    receiver_email = [ 'igor.chaves@maisproxima.com.br', 'Andreia Santos <andreia.santos@maisproxima.com.br>' , 'sheila.gomes@maisproxima.com.br' , 'anderson.souza@maisproxima.com.br', 'Artur Nucci Ferrari <ti@maisproxima.com.br>', 'Tatiana Mello <tatiana.mello@maisproxima.com.br>']
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

