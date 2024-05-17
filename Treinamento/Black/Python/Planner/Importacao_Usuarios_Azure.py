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



url = "https://graph.microsoft.com/beta/users"

payload = {}
headers = {
  'SdkVersion': 'postman-graph/v1.0',
  'Authorization': token
}

response = requests.request("GET", url, headers=headers, data = payload)

j = simplejson.loads(response.text)

cursor.execute("TRUNCATE TABLE PowerBIV2..VB_Usuarios_Azure")

cont = 0
for items in j['value']:
    user_id = "'" + str(j["value"][cont]['id']) + "'"
    user_displayName = "'" + str(j["value"][cont]['displayName']) + "'"
    user_givenName = "'" + str(j["value"][cont]['givenName']) + "'"
    user_jobTitle = "'" + str(j["value"][cont]['jobTitle']) + "'"
    user_mail = "'" + str(j["value"][cont]['mail']) + "'"
    user_mobilePhone = "'" + str(j["value"][cont]['mobilePhone']) + "'"
    user_officeLocation = "'" + str(j["value"][cont]['officeLocation']) + "'"
    user_preferredLanguage = "'" + str(j["value"][cont]['preferredLanguage']) + "'"
    user_surname = "'" + str(j["value"][cont]['surname']) + "'"
    user_userPrincipalName = "'" + str(j["value"][cont]['userPrincipalName']) + "'"

    query = "INSERT INTO PowerBIv2..VB_Usuarios_Azure VALUES(" + user_displayName + "," + user_givenName.replace("None","") + "," + user_jobTitle.replace("None","") + "," + user_mail + "," + user_mobilePhone.replace("None","") + "," + user_officeLocation.replace("None","") + "," + user_preferredLanguage.replace("None","") + "," + user_surname.replace("None","") + "," + user_userPrincipalName.replace("None","") + "," + user_id + ", NULL, NULL,NULL)"
    cursor.execute(query)
    cont = cont + 1

cursor.execute("EXEC PROC_RUN_ROTINAS 'Importar_Usuarios_Azure'")
cursor.execute("EXEC Check_Rotinas_Python 'Importar_Usuários_Planner'")



