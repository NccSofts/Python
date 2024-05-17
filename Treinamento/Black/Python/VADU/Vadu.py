import pandas as pd
import sqlalchemy
import pyodbc
import requests
import simplejson
import ftplib
import os
import json
import urllib.request
from zipfile import ZipFile

#from sqlalchemy import create_engine

# Cria variaveis para filtro automatico das datas
from datetime import timedelta, date



def left(s, amount):
    return s[:amount]

def right(s, amount):
    return s[-amount:]

def mid(s, offset, amount):
    return s[offset:offset+amount]


def createFolder(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print('Error: Criando diretório. ' + directory)

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



def grabFile(fname):
    localfile = open('C:/Python/TEMP/' + fname, 'wb')
    ftp.retrbinary('RETR ' + fname, localfile.write, 1024)
    localfile.close()

def placeFile(fname):
    ftp.storbinary('STOR ' + fname, open('/old/' + fname, 'rb'))



hoje = date.today()
datai = date.today()
dataf = datai - timedelta(15)
dataini = str(datai)
datafinal = str(dataf)


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

# Conexao com o servidor SQL
engine = sqlalchemy.create_engine(sql_string)
engine.connect()


# CONEXÃO COM FTP
server = '52.67.196.1'
user = 'vadu'
pswd = '8P6na#lmIs'



# PEGAR DATA DA ULTIMA IMPORTAÇÃO VADU
url = "https://www.vadu.com.br/vadu.dll/ServicoGrupoMonitoramento/DataUltimoSerasaImportado"

payload = {}
headers = {
  'Authorization': 'Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiJWYWR1IiwidXNyIjozNzE3LCJlbWwiOiJtb3RvckBtYWlzcHJveGltYS5jb20uYnIiLCJlbXAiOjI1NTkxMDZ9.xgH8rQS7ERUHZf9c8ft2-Z9EPykuYvFfGckMcAXDhiQ',
  'Cookie': 'VADUSESSIONID=%7B81511C6D-6494-4F30-9223-CE86EC7C0926%7D; _QRW=687860786'
}

response = requests.request("GET", url, headers=headers, data = payload)

# print(response.text.encode('utf8'))

obj = simplejson.loads(response.text)

data_ultimo_serasa_importado = obj["DataUltimoSerasaImportado"]

print(data_ultimo_serasa_importado)


#Open ftp connection
ftp = ftplib.FTP(server, user, pswd)
ftp.encoding = 'utf-8'
print(ftp.getwelcome())

ftp.cwd("/")

temp_directory = 'C:/Python/TEMP/'
createFolder(temp_directory)

# TRABALHANDO ARQUIVOS NO FTP
try:
    lista_diretorio = ftp.nlst()

    filename = "VIEW_VADU_2_CedenteResumo_" + str(hoje) + ".zip"
    arquivo1 = 'VIEW_VADU_2_CedenteResumo.zip'

    try:
        ftp.delete("/old/" + filename)
        print('Arquivo ' + filename + ' deletado da pasta old.')
    except:
        print('Arquivo ' + filename + ' não encontrado na pasta old.')

    lin = 0
    for item in lista_diretorio:
        ftp_file = lista_diretorio[lin]
        teste = ftp_file == arquivo1

        if ftp_file == arquivo1:
            try:
                print('--> Fazendo download do arquivo ' + arquivo1)
                grabFile(arquivo1)

                try:
                    print('Ultima Importação: ' + data_ultimo_serasa_importado)
                    print('')
                    print('--> Enviando arquivo ' + arquivo1 + ' para VADU')
                    url_vadu_envio = 'https://www.vadu.com.br/vadu.dll/ServicoGrupoMonitoramento/UploadZipCedenteResumo2'
                    headers = {
                        'Authorization': 'Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiJWYWR1IiwidXNyIjozNzE3LCJlbWwiOiJtb3RvckBtYWlzcHJveGltYS5jb20uYnIiLCJlbXAiOjI1NTkxMDZ9.xgH8rQS7ERUHZf9c8ft2-Z9EPykuYvFfGckMcAXDhiQ',
                        'Content-Type': 'application/zip'
                    }
                    payload = open(temp_directory + arquivo1, 'rb')
                    response2 = requests.request("POST", url_vadu_envio, headers=headers, data=payload)
                    print('--> Resposta servidor VADU: ' + str(response2.text))
                    payload.close()
                    resposta = str(response2.status_code)
                    if resposta != '500':
                        print('--> Movendo arquivo ' + arquivo1 + ' para a pasta old e renomeando para ' + filename)
                        ftp.rename(arquivo1, "/old/" + filename)
                except:
                    print('Erro ao tentar enviar arquivo')

            except:
                print('Arquivo ' + arquivo1 + ', não encontrado')

        lin = lin + 1



    # ENVIANDO ARQUIVO
    filename = "VIEW_VADU_netConsultaSerasa_" + str(hoje) + ".zip"
    arquivo1 = 'VIEW_VADU_netConsultaSerasa.zip'

    try:
        ftp.delete("/old/" + filename)
        print('Arquivo ' + filename + ' deletado da pasta old.')
    except:
        print('Arquivo ' + filename + ' não encontrado na pasta old.')

    lin = 0
    for item in lista_diretorio:
        ftp_file = lista_diretorio[lin]
        teste = ftp_file == arquivo1

        if ftp_file == arquivo1:
            try:
                print('--> Fazendo download do arquivo ' + arquivo1)
                grabFile(arquivo1)

                try:
                    print('--> Enviando arquivo ' + arquivo1 + ' para VADU')
                    url_vadu_envio = 'https://www.vadu.com.br/vadu.dll/ServicoGrupoMonitoramento/UploadZipConsultaSerasa'
                    headers = {
                        'Authorization': 'Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiJWYWR1IiwidXNyIjozNzE3LCJlbWwiOiJtb3RvckBtYWlzcHJveGltYS5jb20uYnIiLCJlbXAiOjI1NTkxMDZ9.xgH8rQS7ERUHZf9c8ft2-Z9EPykuYvFfGckMcAXDhiQ',
                        'Content-Type': 'application/zip'
                    }
                    payload = open(temp_directory + arquivo1, 'rb')
                    response2 = requests.request("POST", url_vadu_envio, headers=headers, data=payload)

                    payload.close()
                    resposta2 = str(response2.status_code)
                    if resposta2 != '500':
                        print('--> Resposta servidor VADU: ' + str(response2.text))
                        print('--> Movendo arquivo ' + arquivo1 + ' para a pasta old e renomeando para ' + filename)
                        ftp.rename(arquivo1, "/old/" + filename)
                    else:
                        print('Erro ao tentar enviar arquivo')
                except:
                    print('Erro ao tentar enviar arquivo')

            except:
                print('Arquivo ' + arquivo1 + ', não encontrado')

        lin = lin + 1




# # # ENVIANDO ARQUIVO
# try:
#
#     lista_diretorio = ftp.nlst()
#
#     filename = "VIEW_VADU_netConsultaSerasa_" + str(hoje) + ".zip"
#     arquivo1 = 'VIEW_VADU_netConsultaSerasa.zip'
#
#     try:
#         ftp.delete("/old/" + filename)
#         print('Arquivo ' + filename + ' deletado da pasta old.')
#     except:
#         print('Arquivo ' + filename + ' não encontrado na pasta old.')
#
#     try:
#         print('--> Fazendo download do arquivo ' + arquivo1)
#         grabFile(arquivo1)
#     except:
#         print('Arquivo ' + arquivo1 + ', não encontrado')
#
#     try:
#         # print('Ultima Importação: ' + data_ultimo_serasa_importado)
#         # print('')
#         print('--> Enviando arquivo para VADU: '+ arquivo1)
#         url_vadu_envio = 'https://www.vadu.com.br/vadu.dll/ServicoGrupoMonitoramento/UploadZipConsultaSerasa'
#         headers = {
#             'Authorization': 'Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpc3MiOiJWYWR1IiwidXNyIjozNzE3LCJlbWwiOiJtb3RvckBtYWlzcHJveGltYS5jb20uYnIiLCJlbXAiOjI1NTkxMDZ9.xgH8rQS7ERUHZf9c8ft2-Z9EPykuYvFfGckMcAXDhiQ',
#             'Content-Type': 'application/zip'
#         }
#         payload = open(temp_directory + arquivo1, 'rb')
#         response2 = requests.request("POST", url_vadu_envio, headers=headers, data=payload)
#         print(response2.text)
#         payload.close()
#         resposta = str(response2.status_code)
#         if resposta != '500':
#             print('--> Movendo arquivo ' + arquivo1 + ' para a pasta old e renomeando para ' + filename)
#             ftp.rename("VIEW_VADU_netConsultaSerasa.zip", "/old/" + filename)
#     except:
#         print('Erro ao tentar enviar arquivo')





except ftplib.error_perm as resp:
    if str(resp) == "550 No files found":
        print("--> Não existem arquivos na pasta")
    else:
        raise

ftp.quit()

