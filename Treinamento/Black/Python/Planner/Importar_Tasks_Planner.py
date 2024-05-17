import requests
import simplejson
import pandas as pd
import pyodbc
import sqlalchemy
from datetime import timedelta, date


def left(s, amount):
    return s[:amount]

def right(s, amount):
    return s[-amount:]

def mid(s, offset, amount):
    return s[offset:offset+amount]

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

# Cria tabela temporária JSON
colunas = ['planId', 'bucketId', 'title', 'orderHint', 'assigneePriority', 'percentComplete', 'startDateTime', 'createdDateTime',
        'dueDateTime','completedDateTime', 'completedBy','priority', 'id_task', 'creator_id', 'creator_name', 'user_resp', 'user_resp_name']
lst = []


def user_token_planner():
    user = "logistica@maisproxima.com.br"
    senha = "Leonardo080200"

    try:
        url = "https://login.microsoftonline.com/b6d1ac3f-3de8-4a1a-853c-bc0081529668/oauth2/v2.0/token"


        payload = "grant_type=password&client_id=63552403-7524-4a03-9396-ce5447045e26&client_secret=AGD6.642WU%7E2yq55_kTRU.wRkj3q-6up-y&scope=https%3A//graph.microsoft.com/.default&userName=" + user + "&password=" + senha
        # payload = 'grant_type=password&client_id=63552403-7524-4a03-9396-ce5447045e26&client_secret=AGD6.642WU%7E2yq55_kTRU.wRkj3q-6up-y&scope=https%3A//graph.microsoft.com/.default&userName=logistica@maisproxima.com.br&password=Leonardo080200'
        headers = {
            'Content-Type': 'application/x-www-form-urlencoded',
            'SdkVersion': 'postman-graph/v1.0',
            'Cookie': 'fpc=AsiXDLklOwlPurlllpGiMPbuKy53AQAAAGcckdYOAAAA; x-ms-gateway-slice=prod; stsservicecookie=ests'
        }

        response = requests.request("POST", url, headers=headers, data = payload)
        total_caracteres = len(response.text)-2
        posicao_token = response.text.find('"access_token":"')+16
        user_token = response.text[posicao_token:total_caracteres]
        print("Token de Usuário Planner OK")
        return user_token

    except:
        print("Erro ao tentar obter token de usuário!!!")

token = 'Bearer ' + user_token_planner()


# Lista de Planos
df = pd.read_sql_table('VB_Planner_Plans', engine)
lin = 0

for row in df.itertuples():

    plan_id = str(df['plan_id'].values[lin])
    # plan_id = '5P8LRhhRDUiwphfC3oUYNmUAFaiZ'
    plan_name = df['title_plan'].values[lin]

    print('Plano de Trabalho' + plan_name)

    url = "https://graph.microsoft.com/beta/planner/plans/" + plan_id + "/tasks"

    payload = {}
    headers = {
        'Content-Type': 'application/json',
        'SdkVersion': 'postman-graph/v1.0',
        'Authorization': token
    }

    response = requests.request("GET", url, headers=headers, data=payload)
    j = simplejson.loads(response.text)
    verify = response.text.find("You do not have the required permissions to access this item")
    total_registros = len(j['value'])


    cont = 0


    for items in j["value"]:

        tarefa = j["value"][cont]['title']
        xxx = j["value"][cont]['assignments']

        print('------> Tarefa: ',tarefa)

        for responsavel in xxx.keys():

            task_planId = str(j["value"][cont]['planId'])
            task_bucketId = str(j["value"][cont]['bucketId'])
            task_title = str(j["value"][cont]['title'])
            task_orderHint = str(j["value"][cont]['orderHint'])
            task_assigneePriority = str(j["value"][cont]['assigneePriority'])
            task_percentComplete = str(j["value"][cont]['percentComplete'])
            task_startDateTime = left(str(j["value"][cont]['startDateTime']), 10)
            task_createdDateTime = left(str(j["value"][cont]['createdDateTime']), 10)
            task_dueDateTime = left(str(j["value"][cont]['dueDateTime']), 10)
            task_completedDateTime = left(str(j["value"][cont]['completedDateTime']), 10)

            try:
                task_completedBy = str(j["value"][cont]['completedBy']['user']['id'])
            except:
                task_completedBy = ''

            task_priority = str(j["value"][cont]['priority'])
            task_id = str(j["value"][cont]['id'])

            try:
                task_creator_id = str(j["value"][cont]['createdBy']['user']['id'])
                task_creator_name = str(j["value"][cont]['createdBy']['user']['displayName'])
            except:
                task_creator_id = ''
                task_creator_name = ''

            task_creator_name = task_creator_name.replace("None", '')

            task_responsavelId = responsavel
            task_responsavel_nome = ''

            lst.append([plan_id, task_bucketId, task_title, task_orderHint, task_assigneePriority,
                        task_percentComplete, task_startDateTime, task_createdDateTime, task_dueDateTime,
                        task_completedDateTime, task_completedBy,
                        task_priority, task_id, task_creator_id, task_creator_name, task_responsavelId, task_responsavel_nome])

        cont = cont + 1

    lin = lin + 1

print('')
print('')
print("Salvando Tarefas na tabela temporária")
df2 = pd.DataFrame.from_records(lst, columns=colunas)
df2.to_sql(
    name='VB_Planner_Tasks_TEMP',  # database table name
    con=engine,
    if_exists='replace',
    index=False,
)

""

cursor.execute("EXEC Base_Tarefas_Planner")
cursor.execute("EXEC Check_Rotinas_Python 'Importar_Tasks_Planner'")

print('')
print("Importação finalizada")