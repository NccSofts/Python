import imaplib
import email
import os
from datetime import date
import re
import Email_Conciliacao_FIDC
import sys
sys.path.append('C:/Python/Database/')
import database as db
import pandas as pd
import base64
base64.MAXBINSIZE = 65536

# Definir conexao 0 = PowerBiV2 / 1 = Datawarehouse
cursor = db.sql_conn('PowerBiv2')
engine = db.sql_engine('PowerBiv2')



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
path_dir = ''

# attatchment_dir = 'C:/CTE_Importados_Email/'
# user = 'cte@maisproxima.com.br'
# password = 'MP#qwert1234'
# imap_url = 'Outlook.office365.com'



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
        global file_Name
        global data_arquivo
        global nome_arquivo

        fileName = part.get_filename()

        data_arquivo = fileName[17:25]
        nome_arquivo = str(data_arquivo +'_Estoque_Fundo_FIDC.xlsx')

        if fileName is not None:
            if fileName.find('.xls') >= 0:
                if bool(fileName):
                    createFolder(path_dir + 'XLS/')
                    # filePath = os.path.join(path_dir + 'XLS/', fileName)
                    global filePath
                    filePath = os.path.join(path_dir + 'XLS/', nome_arquivo)
                    with open(filePath, 'wb') as f:
                        f.write(part.get_payload(decode=True))
                        print('Arquivo ' + nome_arquivo + ' salvo...')


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
        n = n + 1
        typ, data = con.fetch(num,'(RFC822)')

        for response_part in data:
            if isinstance(response_part, tuple):
                original = email.message_from_bytes(response_part[1])
                from_e = str(original['Return-Path'])
                assunto_e = str(original['Subject'])
                print('De: ' + from_e)
                # print('Assunto: ' + assunto_e)
                data_email = email.utils.parsedate_tz(original['Date'])
                data_atual = str(date.today())
                data_email2 = str(data_email[0]) + '-' + str("{:02d}".format(data_email[1])) + '-' + str("{:02d}".format(data_email[2]))
                # print(data_email)
                print('Recebido em: ' + data_email2)
                typ, data = con.store(num,'+FLAGS','\\Seen')
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

                path_dir = attatchment_dir + data_email2 + '/' + sender + '/'
                print('Gravando anexos na pasta: ' + path_dir)
                createFolder(path_dir)
                get_attatchment(original)
                vazio = str(is_dir_empty(path_dir))
                if vazio == 'True':
                    os.rmdir(path_dir)
                    print('Deletando pasta ' + path_dir + ' por estar vazia')

                df = pd.read_excel(filePath)

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
                                          ,'j' --[PES_TIPO_PESSOA]
                                          ,[CedenteCnpjCpf]
                                          ,[CedenteNome]
                                          ,[NotaPdd]
                                          ,'j' --[SAC_TIPO_PESSOA]
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
                                          ,CAST([DataPosicao] AS datetime)
                                          ,CAST([DataProrrogacao] AS datetime)
                                          ,CAST([DataOcorrenciaProrrogacao] AS datetime)
                                          ,[Estorno PDD Vencido]
                                          ,CAST(Data_Arquivo AS DATE) Data_Arquivo
                                      FROM [PowerBIv2].[dbo].[Tabela_Base_DL_Nexus_TEMP]
                                      """
                cursor.execute(insert_arquivo)
                print('Enviando Email')
                Email_Conciliacao_FIDC.Email_FIDC()
                cursor.execute('EXEC PowerBIv2..PROC_Analise_Historica_FIDC')

print('')
print('Final das importações')



