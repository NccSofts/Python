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
from datetime import timedelta, date

diasemana = datetime.datetime.today().weekday()

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

if diasemana < 5:

    linha_vazia = '<p>&nbsp;</p>'

    query = """\
    SELECT 
        Nome Vendedor,
        Canal,
        IIF(MetaFormatada IS NULL, '0', MetaFormatada) Meta,
        IIF(VendaFormatada IS NULL, '0', VendaFormatada) Venda,
        FORMAT(P_Meta,'0.#%') [%Meta],
        Ticket_Medio_Formatado TicketMedio,
        FORMAT(COALESCE((NULLIF(VendaMes,0) / D2.DiasUteis) * D.DiasUteis,0), 'C', 'PT-BR') ProjVenda,
        FORMAT(COALESCE(((NULLIF(VendaMes,0) / D2.DiasUteis) * D.DiasUteis) / NULLIF(Meta,0),0), '0.#%') [%Projeçao],
        GrupoCarteira,
        QtdContEfetivos [Cont.Efetivos],
        FORMAT(P_Contatos,'0.#%') [%Contatos],
        GrupoCompra,
        FORMAT(COALESCE(P_Cont_Ativos, 0),'0.##%') [%Conv.Ativos],
        FORMAT(COALESCE(P_Cont_Inativos, 0),'0.##%') [%Conv.Inativos],
        QtdContOutros [Cont.Outros],
        QtdContTentativas [Cont.Tentativas],
        QtdLigEfetivas [Lig.Efetivas],
        QtdLigOutros [Lig.Outros],
        QtdLigTentativas [Lig.tentativas]
    FROM [PowerBIv2].[dbo].[Carteira_Clientes_Resumo_Email] C
    INNER JOIN (
                SELECT 
                    Ano,
                    Mes,
                    COUNT(Util) DiasUteis
                FROM PowerBIv2..Tabela_Datas
                WHERE Util = 'S' AND Ano = FORMAT(GETDATE(),'yyyy') AND Mes = FORMAT(GETDATE(),'MM')
                GROUP BY Ano, Mes
                ) D ON D.Ano = FORMAT(GETDATE(),'yyyy') AND D.Mes = FORMAT(GETDATE(),'MM')
    INNER JOIN (
                SELECT 
                    Ano,
                    Mes,
                    COUNT(Util) DiasUteis
                FROM PowerBIv2..Tabela_Datas
                WHERE Util = 'S' AND Data BETWEEN DATEADD(DAY, 1, EOMONTH(GETDATE(), -1)) AND GETDATE()-1
                GROUP BY Ano, Mes
                ) D2 ON D2.Ano = FORMAT(GETDATE(),'yyyy') AND D2.Mes = FORMAT(GETDATE(),'MM')
    ORDER BY Canal, Nome asc
    """

    df = pd.read_sql_query(query, engine)
    df.to_html('C:\Python\Gauge\Lista_Vendedores.html')


    query = """\
    SELECT 
        Canal,
        FORMAT(COALESCE(SUM(Meta), 0), 'C', 'PT-BR') Meta,
        FORMAT(COALESCE(SUM(VendaMes), 0), 'C', 'PT-BR') Venda,
        FORMAT(COALESCE(SUM(VendaMes) / NULLIF(SUM(Meta),0),0), '##.#%') [%Meta],
        FORMAT(COALESCE((SUM(VendaMes) / NULLIF(D2.DiasUteis,0)) * D.DiasUteis,0), 'C', 'PT-BR') ProjeçaoVenda,
        FORMAT(COALESCE(((SUM(VendaMes) / NULLIF(D2.DiasUteis,0)) * D.DiasUteis) / NULLIF(SUM(Meta),0),null), '##.#%') [%Projeçao]
		--FORMAT(COALESCE(SUM(GrupoCompra) / SUM(GrupoCarteira), 0),'##%') [%Conv.Carteira]
    
    FROM [PowerBIv2].[dbo].[Carteira_Clientes_Resumo_Email] C
    INNER JOIN (
                SELECT 
                    Ano,
                    Mes,
                    COUNT(Util) DiasUteis
                FROM PowerBIv2..Tabela_Datas
                WHERE Util = 'S' AND Ano = FORMAT(GETDATE(),'yyyy') AND Mes = FORMAT(GETDATE(),'MM')
                GROUP BY Ano, Mes
                ) D ON D.Ano = FORMAT(GETDATE(),'yyyy') AND D.Mes = FORMAT(GETDATE(),'MM')
    INNER JOIN (
                SELECT 
                    Ano,
                    Mes,
                    COUNT(Util) DiasUteis
                FROM PowerBIv2..Tabela_Datas
                WHERE Util = 'S' AND Data BETWEEN DATEADD(DAY, 1, EOMONTH(GETDATE(), -1)) AND GETDATE()-1
                GROUP BY Ano, Mes
                ) D2 ON D2.Ano = FORMAT(GETDATE(),'yyyy') AND D2.Mes = FORMAT(GETDATE(),'MM')
    GROUP BY Canal, D.DiasUteis, D2.DiasUteis, C.Canal
    
    UNION ALL
    
    SELECT 
        'TOTAL',
        FORMAT(COALESCE(SUM(Meta), 0), 'C', 'PT-BR') Meta,
        FORMAT(COALESCE(SUM(VendaMes), 0), 'C', 'PT-BR') Venda,
        FORMAT(COALESCE(SUM(VendaMes) / NULLIF(SUM(Meta),0),0), '##.#%') [%Meta],
        FORMAT(COALESCE((SUM(VendaMes) / NULLIF(D2.DiasUteis,0)) * D.DiasUteis,0), 'C', 'PT-BR') ProjeçaoVenda,
        FORMAT(COALESCE(((SUM(VendaMes) / NULLIF(D2.DiasUteis,0)) * D.DiasUteis) / NULLIF(SUM(Meta),0), NULL), '##.#%') [%Projeçao]
        --FORMAT(COALESCE(SUM(GrupoCompra) / SUM(GrupoCarteira), 0),'##%') [%Conv.Carteira]
    
    FROM [PowerBIv2].[dbo].[Carteira_Clientes_Resumo_Email] C
    INNER JOIN (
                SELECT 
                    Ano,
                    Mes,
                    COUNT(Util) DiasUteis
                FROM PowerBIv2..Tabela_Datas
                WHERE Util = 'S' AND Ano = FORMAT(GETDATE(),'yyyy') AND Mes = FORMAT(GETDATE(),'MM')
                GROUP BY Ano, Mes
                ) D ON D.Ano = FORMAT(GETDATE(),'yyyy') AND D.Mes = FORMAT(GETDATE(),'MM')
    INNER JOIN (
                SELECT 
                    Ano,
                    Mes,
                    COUNT(Util) DiasUteis
                FROM PowerBIv2..Tabela_Datas
                WHERE Util = 'S' AND Data BETWEEN DATEADD(DAY, 1, EOMONTH(GETDATE(), -1)) AND GETDATE()-1
                GROUP BY Ano, Mes
                ) D2 ON D2.Ano = FORMAT(GETDATE(),'yyyy') AND D2.Mes = FORMAT(GETDATE(),'MM')
    GROUP BY D.DiasUteis, D2.DiasUteis
    """

    style_tabela = """\
    <style>
    table {
      border-collapse: collapse;
      width: 100%;
    }
    
    th, td {
      text-align: left;
      padding: 8px;
    }
    
    tr:nth-child(even) {background-color: #f2f2f2;}
    </style>
    """
    df = pd.read_sql_query(query, engine)
    df.to_html('C:\Python\Gauge\Total_Venda.html')

    #Concatenando códigos HTML num unico arquivo

    # Cabeçalho do Email
    # with open("C:\Python\Gauge\Resumo_Vendas.html", "w") as file:
    #     file.write(style_tabela)

    tabela_sql = open('C:\Python\Gauge\Lista_Vendedores.html', 'r')
    html = tabela_sql.read()
    html = html.replace('<tr style="text-align: right;">','<tr style="text-align: center;">')
    html = html.replace('<td>','<td style="text-align: center;">')

    with open("C:\Python\Gauge\Resumo_Vendas.html", "w") as file:
        file.write(html)

    with open("C:\Python\Gauge\Resumo_Vendas.html", "a") as file:
        file.write(linha_vazia)

    tabela_sql = open('C:\Python\Gauge\Total_Venda.html', 'r')
    html = tabela_sql.read()
    with open("C:\Python\Gauge\Resumo_Vendas.html", "a") as file:
        file.write(html)


    corpo = codecs.open("C:\Python\Gauge\Resumo_Vendas.html", 'r')
    html = corpo.read()
    html = html.replace('<tr style="text-align: right;">','<tr style="text-align: center;">')
    html = html.replace('<td>','<td style="text-align: center;">')

    # ENVIA EMAIL
    diasemana = date.today().weekday()
    data_atual = date.today()
    data_pt_br = data_atual.strftime('%d/%m/%Y')
    # sender_email = "credito@maisproxima.com.br"
    user_login = 'bi@maisproxima.com'
    # receiver_email = ['anderson.souza@maisproxima.com.br']
    # receiver_email = ['logistica@maisproxima.com.br']
    receiver_email = ['carlos.souza@maisproxima.com.br','simone.dario@maisproxima.com.br','tatiana.mello@maisproxima.com.br']
    # password = input("Type your password and press enter:")
    # password = 'MP#qwert1234'
    password = '+Proxima2019'

    titulo_email = 'PERFORMANCE DE VENDAS - ' + data_pt_br

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