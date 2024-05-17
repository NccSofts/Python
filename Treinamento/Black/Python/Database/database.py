import os
import pyodbc
import sqlalchemy
import pandas as pd
import datetime
from datetime import timedelta, date
import pandas as pd
import time
import requests
import simplejson

# CRIANDO CONEXÃO COM OS BANCOS SQL
def sql_conn(opcao_bd):

    # Definir conexao 0 = PowerBiV2 / 1 = Datawarehouse

    server = '52.67.126.131'
    database = opcao_bd
    username = 'python'
    password = '?oST#hECh7*EqlFR3S?i'
    
    sql_string = 'mssql+pyodbc://' + username + ':' + password + '@' + server + '/' + database + '?driver=SQL+SERVER'

    return pyodbc.connect(
        'DRIVER={SQL SERVER};SERVER=' + server + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password,
        autocommit=True)
    cursor = cnxn.cursor()


def sql_engine(opcao_bd):

    # Definir conexao 0 = PowerBiV2 / 1 = Datawarehouse

    server = '52.67.126.131'
    database = opcao_bd
    username = 'python'
    password = '?oST#hECh7*EqlFR3S?i'

    sql_string = 'mssql+pyodbc://' + username + ':' + password + '@' + server + '/' + database + '?driver=SQL+SERVER'

    server = '52.67.126.131'
    database = opcao_bd
    username = 'python'
    password = '?oST#hECh7*EqlFR3S?i'

    # Conexao com o servidor SQL
    return sqlalchemy.create_engine(sql_string)
    engine.connect()


# FUNÇÕES PARA TRABALHAR COM TEXTOS
def left(s, amount):
    return s[:amount]

def right(s, amount):
    return s[-amount:]

def mid(s, offset, amount):
    return s[offset:offset+amount]

def hoje():
    return date.today().strftime('%d/%m/%Y')

def hora_atual():
    return str(datetime.datetime.today().time())

def diadasemana():
    return datetime.datetime.today().weekday()

def createFolder(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
            return print('Diretório ' + directory + ' criado.')
    except OSError:
        return print('Erro ao criar diretório ' + directory)

def token_i9():

    # OBTENDO CHAVE DE ACESSO A API DO I9
    print('OBTENDO TOKEN DE ACESSO A API')
    print('-----------------------------')
    url = "https://api.ifieldlogistics.com/v1/auth"
    # url = "https://sandbox-api.ifieldlogistics.com/v1/auth/"
    payload = {
        'clientTokenApi': "TlcoSDRrOXQlSTI4MDIpNjAk.eyJ1c2VybmFtZSI6ImFuZGVyc29uLnNvdXphIiwicGFzc3dvcmQiOiJld2lBZUV3NlRkNVhzUTNWR0pcL3cwQT09In0.SuQZzmPKcO3YlkvVw20h2eHxlIq6F17IKfmnYgH3njg"}
    headers = {
        'content-type': "multipart/form-data; boundary=----WebKitFormBoundary7MA4YWxkTrZu0gW",
        'Content-Type': "application/x-www-form-urlencoded",
        'Cache-Control': "no-cache",
        'Postman-Token': "e2061448-9b6c-4a08-b20b-3a5d1b21d41f"
    }
    response = requests.request("POST", url, data=payload, headers=headers)
    numcaract = len(response.text) - 14
    tokenAPI = 'Bearer ' + mid(response.text, 13, numcaract)
    tokenAPI = tokenAPI.replace('"', '')
    return tokenAPI