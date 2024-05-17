import imaplib
import email
import os
from datetime import date
from googleapiclient.discovery import build
# from googleapiclient.discovery import build
# pip install --upgrade google-api-python-client
import re
import subprocess
import Email_Conciliacao_FIDC
import smtplib, ssl
import pandas as pd
import win32com.client
import os
import getpass
from datetime import timedelta, date
import pyodbc
import sqlalchemy
import  numpy as np


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





def change_date_format(dt):
    return re.sub(r'(\d{4})-(\d{1,2})-(\d{1,2})', '\\3-\\2-\\1', dt)


def is_dir_empty(path):
    return next(os.scandir(path), None) is None

def getListOfFiles(dirName):
    # create a list of file and sub directories
    # names in the given directory
    listOfFile = os.listdir(dirName)
    allFiles = list()
    # Iterate over all the entries
    for entry in listOfFile:
        # Create full path
        fullPath = os.path.join(dirName, entry)
        # If entry is a directory then get the list of files in this directory
        if os.path.isdir(fullPath):
            allFiles = allFiles + getListOfFiles(fullPath)
        else:
            allFiles.append(fullPath)

    return allFiles



attatchment_dir = 'C:/Importacoes/'
imap_url = 'imap.gmail.com'
user = 'bi@maisproxima.com'
password = '+Proxima2019'

# imap_url = 'imap.mail.yahoo.com'
# user = 'bi.maisproxima@yahoo.com'
# password = "qhhxzyeptczpdzls"
# path_dir = ''

arquivo = ''

def substring(s, start, end):
    return s[start:end]


def createFolder(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print('Error: Criando diretório. ' + directory)


def get_attatchment(msg):
    for part in msg.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue

        fileName = data_arquivo + '_Estoque_Fundo_FIDC.xlsx'


        if fileName.find('.xls') >= 0:
            if bool(fileName):
                createFolder(path_dir + 'Excel/')
                filePath = os.path.join(path_dir + 'Excel/', fileName)
                with open(filePath, 'wb') as f:
                    f.write(part.get_payload(decode=True))
                    print('Arquivo ' + fileName + ' salvo...')


        # if fileName.find('.xlsx') >= 0:
        #     if bool(fileName):
        #         createFolder(path_dir + 'Excel/')
        #         filePath = os.path.join(path_dir + 'Excel/', fileName)
        #         with open(filePath, 'wb') as f:
        #             f.write(part.get_payload(decode=True))
        #             print('Arquivo ' + fileName + ' salvo...')

        # if fileName.find('.pdf') >= 0:
        #     if bool(fileName):
        #         createFolder(path_dir + 'Excel/')
        #         filePath = os.path.join(path_dir + 'Excel/', fileName)
        #         with open(filePath, 'wb') as f:
        #             f.write(part.get_payload(decode=True))
        #             print('Arquivo ' + fileName + ' salvo...')


fileName = ''


con = imaplib.IMAP4_SSL(imap_url)
con.login(user, password)
con.select('Fundo_Nexus')
n = 0
(retcode, messages) = con.search(None, '(UNSEEN)')
if retcode == 'OK':
    print('IMPORTANDO EMAILS DE bi@maisproxima.com')

    for num in messages[0].split():
        print('')
        print('Processando email... ')
        n=n+1
        typ, data = con.fetch(num,'(RFC822)')
        for response_part in data:
            if isinstance(response_part, tuple):
                original = email.message_from_bytes(response_part[1])
                from_e = str(original['Return-Path'])
                assunto_e = str(original['Subject'])
                teste = assunto_e.find("Envio do arquivo")
                data_arquivo = assunto_e[41:49].replace('.','-')
                fileName = data_arquivo + '_Estoque_Fundo_FIDC.xlsx'
                print('De: ' + from_e)
                # print('Assunto: ' + assunto_e)
                data_email = email.utils.parsedate_tz(original['Date'])
                data_atual = str(date.today())
                data_email2 = str(data_email[0]) + '-' + str("{:02d}".format(data_email[1])) + '-' + str("{:02d}".format(data_email[2]))
                # print(data_email)
                print('Recebido em: ' + data_email2)
                # typ, data = con.store(num,'+FLAGS','\\Seen')
                sender = str(original['Return-Path'])
                sender = sender.replace('<','')
                sender = sender.replace('>','')
                # print(sender)
                # print(sender.find('@'))
                # print(len(sender))
                emailI = len(sender)
                emailF = sender.find('@')+1
                emailT = emailI + (emailF-1)
                # print(emailT)
                # print(substring(sender,emailF,emailT))
                sender = sender[emailF:emailT]

                path_dir = attatchment_dir
                print('Gravando anexos na pasta: ' + path_dir)
                createFolder(path_dir)
                get_attatchment(original)
                vazio = str(is_dir_empty(path_dir))
                if vazio == 'True':
                    os.rmdir(path_dir)
                    print('Deletando pasta ' + path_dir + ' por estar vazia')



                # Salvando dados da Planilha no SQL

                # Data de hoje
                hoje = date.today().strftime('%d/%m/%Y')

                # Busca os dados da planilha para pesquisar nos sites
                # desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
                desktop_path = "C:/Importacoes/Excel/"
                arquivo = fileName
                desktop = desktop_path + arquivo

                # Name file variables
                file_path = desktop_path
                file_name = arquivo

                full_name = os.path.join(file_path, file_name)

                xl_app = win32com.client.Dispatch('Excel.Application')
                pwd = 'NEXUS36'

                xl_wb = xl_app.Workbooks.Open(full_name, False, True, None, pwd)
                xl_app.Visible = False
                xl_sh = xl_wb.Worksheets('RelatorioEstoque')

                # Get last_row
                row_num = 0
                cell_val = ''
                while cell_val != None:
                    row_num += 1
                    cell_val = xl_sh.Cells(row_num, 1).Value
                    # print(row_num, '|', cell_val, type(cell_val))
                last_row = row_num - 1
                # print(last_row)

                # Get last_column
                col_num = 0
                cell_val = ''
                while cell_val != None:
                    col_num += 1
                    cell_val = xl_sh.Cells(1, col_num).Value
                    # print(col_num, '|', cell_val, type(cell_val))
                last_col = col_num - 1
                # print(last_col)

                content = xl_sh.Range(xl_sh.Cells(1, 1), xl_sh.Cells(last_row, last_col)).Value
                list(content)
                df = pd.DataFrame(list(content[1:]), columns=content[0])
                df["DataEmissao"] = df["DataEmissao"].dt.tz_convert(None)
                df["DataAquisicao"] = df["DataAquisicao"].dt.tz_convert(None)
                df["DataVencimento"] = df["DataVencimento"].dt.tz_convert(None)
                df["DataPosicao"] = df["DataPosicao"].dt.tz_convert(None)
                df.head()

                df.to_sql(
                    name='Tabela_Base_DL_Nexus_TEMP',  # database table name
                    con=engine,
                    if_exists='replace',
                    index=False,
                )

                update_sql = """\
                                UPDATE [PowerBIv2].[dbo].[Tabela_Base_DL_Nexus_TEMP]
                                    SET [CedenteCnpjCpf] = REPLACE([CedenteCnpjCpf],'.','') 

                                UPDATE [PowerBIv2].[dbo].[Tabela_Base_DL_Nexus_TEMP]
                                    SET [CedenteCnpjCpf] = REPLACE([CedenteCnpjCpf],'/','')

                                UPDATE [PowerBIv2].[dbo].[Tabela_Base_DL_Nexus_TEMP]
                                    SET [CedenteCnpjCpf] = REPLACE([CedenteCnpjCpf],'-','')


                                --RETIRANDO CARACTERES ESPECIAIS DOS CNPJS
                                UPDATE [PowerBIv2].[dbo].[Tabela_Base_DL_Nexus_TEMP]
                                    SET [SacadoCnpjCpf] = REPLACE([SacadoCnpjCpf],'.','') 

                                UPDATE [PowerBIv2].[dbo].[Tabela_Base_DL_Nexus_TEMP]
                                    SET [SacadoCnpjCpf] = REPLACE([SacadoCnpjCpf],'/','')

                                UPDATE [PowerBIv2].[dbo].[Tabela_Base_DL_Nexus_TEMP]
                                    SET [SacadoCnpjCpf] = REPLACE([SacadoCnpjCpf],'-','')

                                UPDATE [PowerBIv2].[dbo].[Tabela_Base_DL_Nexus_TEMP]
                                    SET [SacadoCnpjCpf] = REPLACE([SacadoCnpjCpf],'-','')

                                UPDATE [PowerBIv2].[dbo].[Tabela_Base_DL_Nexus_TEMP]
                                    SET NumeroTitulo = REPLACE(NumeroTitulo,'-','')

                                UPDATE [PowerBIv2].[dbo].[Tabela_Base_DL_Nexus_TEMP]
                                SET NumeroTitulo = REPLACE(NumeroTitulo,'.','')
                            """

                cursor.execute(update_sql)

                add_field = f"ALTER TABLE PowerBiv2..Tabela_Base_DL_Nexus_TEMP ADD Data_Arquivo varchar(15)"
                cursor.execute(add_field)

                data_arquivo = "'" + data_arquivo + "'"

                update_data = f"UPDATE PowerBiv2..Tabela_Base_DL_Nexus_TEMP SET Data_Arquivo = {data_arquivo}"
                cursor.execute(update_data)

                insert_arquivo = """\
                
                                    DELETE FROM [PowerBIv2].[dbo].[Tabela_Base_Nexus]
                    
                                    WHERE Data_Arquivo = (SELECT MAX(CAST(Data_Arquivo AS DATE)) FROM [PowerBIv2].[dbo].[Tabela_Base_DL_Nexus_TEMP])
                                    
                                    
                                    INSERT INTO [PowerBIv2].[dbo].[Tabela_Base_Nexus]
                                    
                                    
                                    SELECT [Situacao]
                                          ,[PES_TIPO_PESSOA]
                                          ,[CedenteCnpjCpf]
                                          ,[CedenteNome]
                                          ,[NotaPdd]
                                          ,[SAC_TIPO_PESSOA]
                                          ,[SacadoCnpjCpf]
                                          ,[SacadoNome]
                                          ,[IdTituloVx]
                                          ,[TipoAtivo]
                                          ,[DataEmissao]
                                          ,[DataAquisicao]
                                          ,[DataVencimento]
                                          ,[NumeroBoletoBanco]
                                          ,[NumeroTitulo]
                                          ,[CampoChave]
                                          ,[CMC7]
                                          ,[ValorAquisicao]
                                          ,[ValorNominal]
                                          ,[ValorPresente]
                                          ,[PDD NOTA]
                                          ,[PDD REGULAMENTO]
                                          ,[DataPosicao]
                                          ,[DataProrrogacao]
                                          ,[DataOcorrenciaProrrogacao]
                                          ,[Estorno PDD Vencido]
                                          ,CAST(Data_Arquivo AS DATE) Data_Arquivo
                                      FROM [PowerBIv2].[dbo].[Tabela_Base_DL_Nexus_TEMP]
                                      """
                cursor.execute(insert_arquivo)
                xl_wb.Close(False)

                print('')
                print("EXEC PowerBIv2..PROC_Analise_Historica_FIDC")

                comando = 'EXEC PowerBIv2..PROC_Analise_Historica_FIDC'
                cursor.execute(comando)
                # while 1:
                #     q = status_check_cursor.execute("select RunningStatus from PowerBIv2..Acompanhamento_Rotinas_SQL_Python where Script_Procedure = 'Analise Historica FIDC'").fetchone()
                #     if q[0] == 0:
                #         break

                print("")
                print("Enviando Email")

                ultima_data = """\

                                --ATUALIZANDO DATA ULTIMO ARQUIVO
                                UPDATE [PowerBIv2].[dbo].[Analise_Historica_FIDC]
                                    SET Data_Ultimo_Arquivo = 'Data_Ultimo_Arquivo'
                                FROM [PowerBIv2].[dbo].[Analise_Historica_FIDC] A
                                WHERE Data = (SELECT MAX(Data) FROM [PowerBIv2].[dbo].[Analise_Historica_FIDC])
                                
                                """
                cursor.execute(ultima_data)


                Email_Conciliacao_FIDC.Email_FIDC()



else:
    print('Não existem planilhas para serem importadas')

cursor.execute("EXEC POWERBIV2..PROC_RUN_ROTINAS 'Importar planilha FIDC'")


cursor.close()








