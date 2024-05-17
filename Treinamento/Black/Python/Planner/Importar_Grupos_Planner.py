import requests
import simplejson
import pandas as pd
import pyodbc
import sqlalchemy
from datetime import timedelta, date


# Definir conexao 0 = PowerBiV2 / 1 = Datawarehouse
opcao = 0

configuracoes = pd.read_csv("C:/Python/database.cfg", header = 0, delimiter=";")
email_config = pd.read_csv("C:/Python/Email_Config.cfg", header = 0, delimiter=";")

#Config Servidor SQL

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



def user_token_planner():

    try:
        url = "https://login.microsoftonline.com/b6d1ac3f-3de8-4a1a-853c-bc0081529668/oauth2/v2.0/token"

        payload = 'grant_type=password&client_id=63552403-7524-4a03-9396-ce5447045e26&client_secret=AGD6.642WU%7E2yq55_kTRU.wRkj3q-6up-y&scope=https%3A//graph.microsoft.com/.default&userName=logistica@maisproxima.com.br&password=Leonardo080200'
        headers = {
            'Content-Type': 'application/x-www-form-urlencoded',
            'SdkVersion': 'postman-graph/v1.0',
            'Cookie': 'fpc=AsiXDLklOwlPurlllpGiMPbuKy53AQAAAGcckdYOAAAA; x-ms-gateway-slice=prod; stsservicecookie=ests'
        }

        response = requests.request("POST", url, headers=headers, data = payload)
        total_caracteres = len(response.text)-2
        posicao_token = response.text.find('"access_token":"')+16
        user_token = response.text[posicao_token:total_caracteres]
        print("Token de Usuário Planner")
        print(user_token)
        return user_token

    except:
        print("Erro ao tentar obter token de usuário!!!")


token = 'Bearer ' + user_token_planner()


url = "https://graph.microsoft.com/v1.0/groups"

payload = {}
headers = {
  'Content-Type': 'application/json',
  'SdkVersion': 'postman-graph/v1.0',
  'Authorization': token
}

response = requests.request("GET", url, headers=headers, data = payload)

j = simplejson.loads(response.text)

cursor.execute("TRUNCATE TABLE PowerBIV2..VB_Planner_Grupos")

cont = 0
for items in j['value']:
    group_id = "'" + str(j["value"][cont]['id']) + "'"
    group_deletedDateTime = "'" + str(j["value"][cont]['deletedDateTime']) + "'"
    group_createdDateTime = "'" + str(j["value"][cont]['createdDateTime']) + "'"
    group_displayName = "'" + str(j["value"][cont]['displayName']) + "'"
    group_mail = "'" + str(j["value"][cont]['mail']) + "'"
    group_mailNickname = "'" + str(j["value"][cont]['mailNickname']) + "'"
    group_mailEnabled = "'" + str(j["value"][cont]['mailEnabled']) + "'"
    group_classification = "'" + str(j["value"][cont]['classification']) + "'"

    query = "INSERT INTO PowerBIv2..VB_Planner_Grupos VALUES(" + group_id + "," + group_deletedDateTime.replace("None","") + "," + group_createdDateTime.replace("None", "") + "," + group_displayName + "," + group_mail.replace("None", "") + "," + group_mailNickname.replace("None", "") + "," + group_mailEnabled.replace("None", "") + "," + group_classification.replace("None","") + ")"
    cursor.execute(query)

    cont = cont + 1


cursor.execute("EXEC PROC_RUN_ROTINAS 'Importar_Grupos_Azure'")
cursor.execute("EXEC Check_Rotinas_Python 'Importar_Grupos_Planner'")
