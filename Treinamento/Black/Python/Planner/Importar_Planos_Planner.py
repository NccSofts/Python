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

df = pd.read_sql_table('VB_Planner_Grupos', engine)
cursor.execute("TRUNCATE TABLE PowerBIV2..VB_Planner_Plans")

lin = 0
for row in df.itertuples():
    group_id = df['id_group'].values[lin]
    group_displayName = str(df['displayName'].values[lin])


    url = "https://graph.microsoft.com/v1.0/groups/" + group_id + "/planner/plans"

    payload = {}
    headers = {
      'Content-Type': 'application/json',
      'SdkVersion': 'postman-graph/v1.0',
      'Authorization': token
    }

    response = requests.request("GET", url, headers=headers, data = payload)

    verify = response.text.find("You do not have the required permissions to access this item")

    if response.text.find("You do not have the required permissions to access this item") == -1:

        j = simplejson.loads(response.text)
        cont = 0
        cont2 = 0
        for items in j['value']:
            plan_id = "'" + str(j["value"][cont]['id']) + "'"
            plan_createdDateTime = "'" + str(j["value"][cont]['createdDateTime']) + "'"
            plan_owner = "'" + str(j["value"][cont]['owner']) + "'"
            plan_title = "'" + str(j["value"][cont]['title']) + "'"

            try:
                for items in j["value"][cont]['createdBy']:
                    plan_createdBy_id = "'" + str(j["value"][cont]['createdBy']['user']['id']) + "'"

            except:
                print("")

            query = "INSERT INTO PowerBIv2..VB_Planner_Plans VALUES(" + plan_createdDateTime.replace("None", "") + "," + plan_owner.replace("None", "") + "," + plan_title.replace("None", "") + "," + plan_id + "," + "NULL" + "," + plan_createdBy_id.replace("None", "") + ", NULL" + ")"
            cursor.execute(query)

        cont = cont + 1

        print("Importando tarefas do grupo " , group_displayName)
        print("")



    else:
        print("Grupo " + group_displayName + " não possui tarefas.")
        print("")

    lin = lin + 1

query_update = """\

            UPDATE [PowerBIv2].[dbo].[VB_Planner_Plans]
            
                SET creator_name = u.displayName
            
            FROM [PowerBIv2].[dbo].[VB_Planner_Plans] P 
            INNER JOIN VB_Usuarios_Azure U ON U.id = p.id_creator
            
            """
cursor.execute(query_update)
cursor.execute("EXEC Check_Rotinas_Python 'Importar_Planos_Planner'")
