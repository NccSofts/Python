import os
import pandas as pd
import pyodbc
import sqlalchemy
import codecs
import smtplib, ssl
import PIL
from PIL import Image
import matplotlib.pyplot as plt
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


print('Executando Mapa_Venda_Estoque')
cursor.execute("EXEC PowerBiv2..Mapa_Venda_Estoque")
cursor.commit()



if diasemana < 5:

    linha_vazia = '<p>&nbsp;</p>'

    texto = """\

    <!-- #######  YAY, I AM THE SOURCE EDITOR! #########-->
    <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Olá,</strong></h1>
    <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Segue relação de itens com dias de estoque maior que 180 dias.</strong></h1>
    <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Att.,</strong></h1>
    <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Equipe de BI Mais Proxima<br />&nbsp;</strong></h1>

    """

    query = """\
                    SELECT 
                    
                        FORMAT(Data, 'dd/MM/yy') Data,
                        SUM(SaldoVProtheus) Estoque
                    
                    FROM [PowerBIv2].[dbo].[Tabela_Mapa_Venda_Estoque]
                    WHERE CAST(Data AS DATE) >= DATEADD(DAY, -180, CAST(GETDATE() AS date)) 
                    GROUP BY Data
                    ORDER By CAST(Data as date) ASC

    """

    df = pd.read_sql_query(query, engine)
    # df = df.sort_values(by=['Data'], ascending=True)
    df.plot(kind='line', title='Histórico Estoque', x='Data',y='Estoque', linewidth=3, markersize=6, figsize=(10, 4))
    plt.savefig('C:\Python\Graficos\estoque.png', dpi=100)
    plt.legend()




    query = """\
				SELECT
                    FORMAT(SUM(E.SaldoVProtheus),'C','pt-br') SaldoVProtheus
                FROM [PowerBIv2].[dbo].[Tabela_Mapa_Venda_Estoque] E
                WHERE FiltroHoje = 'Hoje' 
				HAVING SUM(E.SaldoVProtheus) > 0                        
				--INNER JOIN (
                --            SELECT TOP 20    
                --                ProdutoX   
                --            FROM [PowerBIv2].[dbo].[Tabela_Mapa_Venda_Estoque]
                --            ORDER BY SaldoVProtheus desc
                --           ) T ON T.ProdutoX = E.ProdutoX
                ORDER BY SUM(E.SaldoVProtheus) DESC
            
                """
    total_df = pd.read_sql_query(query, engine)
    total_geral = str(total_df['SaldoVProtheus'].values)



    with open("C:\Python\Fiscal\Lista_Mapa_Venda.html", "w") as file:
        file.write(texto)


    tab_img1 = '<img src="cid:image1" alt="" width="800" height="300" /></p>'


    with open("C:\Python\Fiscal\Lista_Mapa_Venda.html", "a") as file:
        file.write(tab_img1)

    with open("C:\Python\Fiscal\Lista_Mapa_Venda.html", "a") as file:
        file.write(linha_vazia)


    query = """\
                SELECT
                                                    
                    Fornecedor,
                    FORMAT(SUM(E.SaldoVProtheus),'C','pt-br') SaldoVProtheus
                                                    
                FROM [PowerBIv2].[dbo].[Tabela_Mapa_Venda_Estoque] E  
                WHERE E.SaldoVProtheus > 0  AND FiltroHoje = 'Hoje'   
                --INNER JOIN (
                --            SELECT TOP 20    
                --                ProdutoX   
                --            FROM [PowerBIv2].[dbo].[Tabela_Mapa_Venda_Estoque]
                --            ORDER BY SaldoVProtheus desc
                --           ) T ON T.ProdutoX = E.ProdutoX
                GROUP BY Fornecedor
                ORDER BY SUM(E.SaldoVProtheus) DESC
               
    """
    df = pd.read_sql_query(query, engine)
    # df = df.sort_values(by=['Order'], ascending=False)
    # df = df.drop(columns="Order")

    df.to_html('C:\Python\Fiscal\Lista_Mapa_Venda_Fornecedor.html')

    with open("C:\Python\Fiscal\Lista_Mapa_Venda_Fornecedor.html", "a") as file:
        file.write(linha_vazia)

    # Concatenando códigos HTML num unico arquivo
    tabela_sql = open('C:\Python\Fiscal\Lista_Mapa_Venda_Fornecedor.html', 'r')
    html = tabela_sql.read()

    with open("C:\Python\API_Facility\Lista_Mapa_Venda.html", "a") as file:
        file.write(linha_vazia)

    with open("C:\Python\Fiscal\Lista_Mapa_Venda.html", "a") as file:
        file.write(html)

    with open("C:\Python\API_Facility\Lista_Mapa_Venda.html", "a") as file:
        file.write(linha_vazia)

    query = """\
                SELECT
                                                
                    E.Fornecedor,
                    E.Grupo,
                    E.ProdutoX,
                    E.Produto,
                    E.SaldoQProtheus,
                    FORMAT(E.SaldoVProtheus,'C','pt-br') SaldoVProtheus
                                                
                FROM [PowerBIv2].[dbo].[Tabela_Mapa_Venda_Estoque] E
                WHERE E.SaldoVProtheus > 0  AND FiltroHoje = 'Hoje'
                --INNER JOIN (
                --            SELECT TOP 20    
                --                ProdutoX   
                --            FROM [PowerBIv2].[dbo].[Tabela_Mapa_Venda_Estoque]
                --            ORDER BY SaldoVProtheus desc
                --            ) T ON T.ProdutoX = E.ProdutoX
                
                ORDER BY E.SaldoVProtheus DESC 

    """

    df = pd.read_sql_query(query, engine)
    # df = df.sort_values(by=['Order'], ascending=False)
    # df = df.drop(columns="Order")
    df.to_html('C:\Python\Fiscal\MapaVenda_Produtos.html')

    # Concatenando códigos HTML num unico arquivo
    tabela_sql = open('C:\Python\Fiscal\MapaVenda_Produtos.html', 'r')
    html = tabela_sql.read()

    with open("C:\Python\Fiscal\Lista_Mapa_Venda.html", "a") as file:
        file.write(html)


    corpo = codecs.open("C:\Python\Fiscal\Lista_Mapa_Venda.html", 'r')
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
    receiver_email = ['logistica@maisproxima.com.br','igor.chaves@maisproxima.com.br','simone.dario@maisproxima.com.br','carlos.souza@maisproxima.com.br','tatiana.mello@maisproxima.com.br']
    # receiver_email = ['fiscal@maisproxima.com.br', 'igor.chaves@maisproxima.com.br']
    # password = input("Type your password and press enter:")
    # password = 'MP#qwert1234'
    password = '+Proxima2019'

    titulo_email = 'Estoque com última compra maior de 180d ---> ' + total_geral

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
    fp = open('C:/Python/Graficos/estoque.png', 'rb')
    msgImage = MIMEImage(fp.read())
    fp.close()

    # Define the image's ID as referenced above
    msgImage.add_header('Content-ID', '<image1>')
    message.attach(msgImage)

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