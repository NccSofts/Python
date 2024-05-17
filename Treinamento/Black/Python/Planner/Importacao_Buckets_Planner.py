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

df = pd.read_sql_table('VB_Planner_Plans', engine)

cursor.execute("TRUNCATE TABLE PowerBIV2..VB_Planner_Buckets")

lin = 0
for row in df.itertuples():
    plan_id = df['plan_id'].values[lin]
    plan_name = df['title_plan'].values[lin]

    # BUSCANDO NA API
    url = "https://graph.microsoft.com/beta/planner/plans/" + plan_id + "/buckets"

    payload = {}
    headers = {
        'Content-Type': 'application/json',
        'SdkVersion': 'postman-graph/v1.0',
        'Authorization': token
    }
    response = requests.request("GET", url, headers=headers, data=payload)
    j = simplejson.loads(response.text)

    cont = 0
    for items in j['value']:
        bucket_name = "'" + str(j["value"][cont]['name']) + "'"
        bucket_id = "'" + str(j["value"][cont]['id']) + "'"
        bucket_planid = "'" + str(j["value"][cont]['planId']) + "'"
        bucket_orderhint = "'" + str(j["value"][cont]['orderHint']) + "'"

        print("Importando tarefa "+ bucket_name + "do plano de trabalho " + plan_name )
        query = "INSERT INTO PowerBIV2..VB_Planner_Buckets VALUES ( " + bucket_name + "," + bucket_planid + "," + bucket_orderhint + "," + bucket_id + " )"
        cursor.execute(query)
        cont = cont + 1

    lin = lin + 1


cursor.execute("EXEC Base_Tarefas_Planner")
cursor.execute("EXEC Check_Rotinas_Python 'Importar_Buckets_Planner'")
