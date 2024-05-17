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






def Email_FIDC():
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

    datai = date.today()
    dataf = datai - timedelta(15)
    dataini = str(datai)
    datafinal = str(dataf)

    linha_vazia = '<p>&nbsp;</p>'

    titulo_query = """\
                        SELECT 

                        CONCAT('Conciliacao FIDC vs Protheus - Foram encontradas ', COUNT(NF) , ' NF(s) com diferença somando um total de ' , FORMAT(SUM(Dif), 'C','PT-BR')) Titulo,
                        SUM(Dif) Total,
                        FORMAT((SELECT MAX(Data_Arquivo) FROM [PowerBIv2].[dbo].[Tabela_Base_NEXUS]), 'dd/MM/yyyy') Data_FIDC

                         FROM Conciliacao_FIDC WHERE Dif <> 0
                     """
    df = pd.read_sql_query(titulo_query, engine)
    Total = df['Total'].values[0]
    titulo = df['Titulo'].values[0]
    data_arquivo = df['Data_FIDC'].values[0]

    if Total != 0:
        titulo = df['Titulo'].values[0]
    else:
        titulo = "Conciliacao FIDC vs Protheus - Não foram encontradas NF´s com diferenca"

    pd.options.display.float_format = '{:,.2f}'.format
    df = pd.read_sql('SELECT * FROM Conciliacao_FIDC', engine)

    if diasemana < 5:
        texto = """\

                <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Olá,</strong></h1>
                <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Segue relação de diferenças encontradas na conciliação do FIDC vs PROTHEUS.</strong></h1>
                <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Att.,</strong></h1>
                <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Equipe de BI Mais Proxima<br />&nbsp;</strong></h1>
                """

        with open("C:\Python\Contabilidade\Conciliacao_FIDC.html", "w") as file:
            file.write(texto)

        with open("C:\Python\Contabilidade\Conciliacao_FIDC.html", "a") as file:
            file.write(linha_vazia)

        with open("C:\Python\Contabilidade\Conciliacao_FIDC.html", "a") as file:
            file.write(f'RESUMO CONCILIAÇÃO FIDC vs PROTHEUS - Posicao FIDC {data_arquivo}')

        df_resumo = pd.pivot_table(df, values=['SaldoFIDC', 'Saldo_CR', 'Dif'], aggfunc=np.sum, index=['Cedente'],
                                   margins=True, margins_name='Total')
        df_resumo.to_html('C:\Python\Contabilidade\Resumo_Conciliacao_FIDC.html')

        # Concatenando códigos HTML num unico arquivo
        tabela_sql = open('C:\Python\Contabilidade\Resumo_Conciliacao_FIDC.html', 'r')
        html = tabela_sql.read()
        #
        with open("C:\Python\Contabilidade\Conciliacao_FIDC.html", "a") as file:
            file.write(html)

        with open("C:\Python\Contabilidade\Conciliacao_FIDC.html", "a") as file:
            file.write(linha_vazia)

        df = pd.read_sql('SELECT * FROM Conciliacao_FIDC WHERE Dif <> 0', engine)
        df_resumo = df
        # df_resumo = pd.pivot_table(df, values=['SaldoFIDC', 'Saldo_CR', 'Dif'], aggfunc=np.sum, index=['Cedente', 'Status', 'CNPJ', 'Cliente', 'NF', 'Data Faturamento', 'Data Entrega', 'Banco'], margins=True, margins_name='Total')
        df_resumo.to_html('C:\Python\Contabilidade\Lista_Conciliacao_FIDC.html', index=False)

        with open("C:\Python\Contabilidade\Conciliacao_FIDC.html", "a") as file:
            file.write('LISTA DE NFS COM DIFERENCA')

        # Concatenando códigos HTML num unico arquivo
        tabela_sql = open('C:\Python\Contabilidade\Lista_Conciliacao_FIDC.html', 'r', errors='ignore')
        html = tabela_sql.read()

        with open("C:\Python\Contabilidade\Conciliacao_FIDC.html", "a") as file:
            file.write(html)

        with open("C:\Python\Contabilidade\Conciliacao_FIDC.html", "a") as file:
            file.write(linha_vazia)

        with open("C:\Python\Contabilidade\Conciliacao_FIDC.html", "a") as file:
            file.write('LISTA DE TITULOS NOVOS')

        novos_titulos = """\
                                SELECT 
                                    IdTitulo,
                                    CAST(Pedido AS varchar) Pedido,
                                    Industria,
                                    [CNPJ-CPF],
                                    [Razao Social],
                                    Valor_Nominal_Entrada Valor,
                                    [Data_Emissao_Titulo]

                                FROM Base_Analise_FIDC
                                WHERE Ultima_Data_Arquivo IS NOT NULL AND Valor_Nominal_Entrada <> 0
                                ORDER BY Industria, IdTitulo
                         """
        df = pd.read_sql_query(novos_titulos, engine)

        if len(df) > 0:

            df.to_html('C:\Python\Contabilidade\Lista_NovosTitlos_FIDC.html', index=False)

            tabela_sql = open('C:\Python\Contabilidade\Lista_NovosTitlos_FIDC.html', 'r', errors='ignore')

            html = tabela_sql.read()

            with open("C:\Python\Contabilidade\Conciliacao_FIDC.html", "a") as file:
                file.write(html)
        else:
            with open("C:\Python\Contabilidade\Conciliacao_FIDC.html", "a") as file:
                file.write(' ---> Não existem titulos novos a serem listados.')

        # VERIFICA SE ESXISTEM TITULOS NOVOS A SEREM LISTADOS
        with open("C:\Python\Contabilidade\Conciliacao_FIDC.html", "a") as file:
            file.write(linha_vazia)

        with open("C:\Python\Contabilidade\Conciliacao_FIDC.html", "a") as file:
            file.write('LISTA DE TITULOS BAIXADOS')

        titulos_baixados = """\
                                SELECT 
                                    IdTitulo,
                                    CAST(Pedido AS varchar) Pedido,
                                    Data_Vencimento,
                                    Industria,
                                    [CNPJ-CPF],
                                    [Razao Social],
                                    Valor_Nominal_Baixa Valor,
                                    [Data_Emissao_Titulo]

                                FROM Base_Analise_FIDC
                                WHERE Ultima_Data_Arquivo IS NOT NULL AND Valor_Nominal_Baixa <> 0
                                ORDER BY Industria, Data_Vencimento
                                 """
        df = pd.read_sql_query(titulos_baixados, engine)

        if len(df) > 0:

            df.to_html('C:\Python\Contabilidade\Lista_TitulosBaixados_FIDC.html', index=False)

            tabela_sql = open('C:\Python\Contabilidade\Lista_TitulosBaixados_FIDC.html', 'r', errors='ignore')

            html = tabela_sql.read()

            with open("C:\Python\Contabilidade\Conciliacao_FIDC.html", "a") as file:
                file.write(html)
        else:
            with open("C:\Python\Contabilidade\Conciliacao_FIDC.html", "a") as file:
                file.write(' ---> Não existem títulos baixados a serem listados.')

        # INCLUI TITULOS VENCIDOS NO EMAIL SE EXISTIREM
        with open("C:\Python\Contabilidade\Conciliacao_FIDC.html", "a") as file:
            file.write(linha_vazia)

        with open("C:\Python\Contabilidade\Conciliacao_FIDC.html", "a") as file:
            file.write('LISTA DE TITULOS VENCIDOS')

        titulos_baixados = """\
                                SELECT 
                                    IdTitulo,
                                    CAST(Pedido AS varchar) Pedido,
                                    Industria,
                                    [CNPJ-CPF],
                                    [Razao Social],
                                    Valor_Nominal Valor,
                                    [Data_Emissao_Titulo],
                                    Data_Vencimento,
                                    StatusVencimento
                                
                                FROM Base_Analise_FIDC
                                WHERE Ultima_Data_Arquivo IS NOT NULL AND StatusVencimento = 'Vencido' AND StatusTitulo = ''
                                ORDER BY Industria, Data_Vencimento
                                 """
        df = pd.read_sql_query(titulos_baixados, engine)

        # VERIFICA SE EXISTEM TITULOS VENCIDOS A SEREM MOSTRADOS
        if len(df) > 0:

            df.to_html('C:\Python\Contabilidade\Lista_TitulosVencidos_FIDC.html', index=False)

            tabela_sql = open('C:\Python\Contabilidade\Lista_TitulosVencidos_FIDC.html', 'r', errors='ignore')

            html = tabela_sql.read()

            with open("C:\Python\Contabilidade\Conciliacao_FIDC.html", "a") as file:
                file.write(html)
        else:
            with open("C:\Python\Contabilidade\Conciliacao_FIDC.html", "a") as file:
                file.write(' ---> Não existem títulos vencidos para serem listados.')

        # ABRE HTML PARA ENVIO POR EMAIL
        corpo = codecs.open("C:\Python\Contabilidade\Conciliacao_FIDC.html", 'r')
        html = corpo.read()

        # ENVIA EMAIL
        diasemana = date.today().weekday()
        data_atual = date.today()
        data_pt_br = data_atual.strftime('%d/%m/%Y')
        user_login = 'bi@maisproxima.com'
        password = '+Proxima2019'
        # receiver_email = ['anderson.souza@maisproxima.com.br']
        receiver_email = ['sheila.gomes@maisproxima.com.br', 'igor.chaves@maisproxima.com.br', 'Leonardo Oliveira <leonardo.oliveira@maisproxima.com.br>', 'anderson.souza@maisproxima.com.br', 'tesouraria@maisproxima.com.br', 'credito@maisproxima.com.br', 'ti@maisproxima.com.br']

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

        # mailserver = smtplib.SMTP('smtp.office365.com', 587)
        mailserver = smtplib.SMTP('smtp.gmail.com', 587)
        mailserver.ehlo()
        mailserver.starttls()
        mailserver.login(user_login, password)
        mailserver.sendmail(sender_email, receiver_email, message.as_string())
        mailserver.quit()

        print('Email Enviado com sucesso')

        cursor.execute("EXEC POWERBIV2..PROC_RUN_ROTINAS 'Envia Email FIDC'")

# Email_FIDC()


