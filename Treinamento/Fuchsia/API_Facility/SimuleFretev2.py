import pandas as pd
import sqlalchemy
import pyodbc
import requests
import simplejson

#from sqlalchemy import create_engine

# Cria variaveis para filtro automatico das datas
from datetime import timedelta, date

def left(s, amount):
    return s[:amount]

def right(s, amount):
    return s[-amount:]

def mid(s, offset, amount):
    return s[offset:offset+amount]


datai = date.today()
dataf = datai - timedelta(15)
dataini = str(datai)
datafinal = str(dataf)


# Definir conexao 0 = PowerBiV2 / 1 = Datawarehouse
opcao = 1

configuracoes = pd.read_csv("C:/Python/database.cfg", header = 0, delimiter=";")

server = configuracoes['IP'].values[opcao]
database = configuracoes['Banco'].values[opcao]
username = configuracoes['Usuario'].values[opcao]
password = configuracoes['Senha'].values[opcao]


sql_string = 'mssql+pyodbc://' + username + ':' + password + '@' + server + '/' + database + '?driver=ODBC+Driver+13+for+SQL+Server'

cnxn = pyodbc.connect(
    'DRIVER={ODBC Driver 13 for SQL Server};SERVER=' + server + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password)
cursor = cnxn.cursor()

# Conexao com o servidor SQL
engine = sqlalchemy.create_engine(sql_string)
engine.connect()

# tokenAPI = "TlcoSDRrOXQlSTI4MDIpNjAk.eyJ1c2VybmFtZSI6ImFuZGVyc29uLnNvdXphIiwicGFzc3dvcmQiOiJld2lBZUV3NlRkNVhzUTNWR0pcL3cwQT09In0.SuQZzmPKcO3YlkvVw20h2eHxlIq6F17IKfmnYgH3njg"

# SOLICITANDO NOVA KEY DE ACESSO A API

print('OBTENDO TOKEN DE ACESSO A API')
print('-----------------------------')

url = "https://api.ifieldlogistics.com/v1/auth"
# url = "https://sandbox-api.ifieldlogistics.com/v1/auth/"


payload = {'clientTokenApi': "TlcoSDRrOXQlSTI4MDIpNjAk.eyJ1c2VybmFtZSI6ImFuZGVyc29uLnNvdXphIiwicGFzc3dvcmQiOiJld2lBZUV3NlRkNVhzUTNWR0pcL3cwQT09In0.SuQZzmPKcO3YlkvVw20h2eHxlIq6F17IKfmnYgH3njg"}
headers = {
    'content-type': "multipart/form-data; boundary=----WebKitFormBoundary7MA4YWxkTrZu0gW",
    'Content-Type': "application/x-www-form-urlencoded",
    'Cache-Control': "no-cache",
    'Postman-Token': "e2061448-9b6c-4a08-b20b-3a5d1b21d41f"
    }

response = requests.request("POST", url, data=payload, headers=headers)

numcaract = len(response.text)-14

tokenAPI = 'Bearer '+ mid(response.text,13,numcaract)
tokenAPI = tokenAPI.replace('"','')
print(tokenAPI)
print('')


# SIMULANDO FRETE NO SIMULE FRETE


lin = 0

# Cria tabela temporária JSON
colunas = ['Order', 'ChavePedido', 'Transportadora', 'CNPJTransp', 'ValorFrete', 'Prazo', 'CNPJCliente', 'RazaoSocial', 'MenorPreco','DataConsulta', 'Cubagem','QtdVolumes', 'PesoTotal']
lst = []



print('IMPORTANDO NOVOS PEDIDOS DE VENDA PARA SIMULACAO')
# engine.execute("DELETE FROM DataWareHouse.dbo.Tabela_Simula_Frete WHERE StatusPedido <> 'Faturado' AND StatusPedido <> 'Faturar Parcial'")
cursor.execute("EXEC Pedidos_SimuleFrete")
cursor.commit()
print('')

df = pd.read_sql_table('vBaseSimuleFretev2', engine)
totalreg = str(len(df))

print('SIMULANDO O FRETE DE '+ totalreg + ' PEDIDOS DE VENDA... AGUARDE')
print('----------------------------------------------------------------')
print('')

contador = 1
for row in df.itertuples():

    Pedido = df['ChavePedido'].values[lin]
    RazaoSocial = df['RazaoSocial'].values[lin]
    CEPOrigem = df['CEPOrigem'].values[lin]
    CEPDestino = df['CEPDestino'].values[lin]
    CNPJOrigem = df['CNPJOrigem'].values[lin]
    CNPJDestino = df['CNPJDestino'].values[lin]
    ValorTotal = str(df['ValorTotal'].values[lin])
    TipoCalculo = df['TipoCalculo'].values[lin]
    PesoTotal = str(df['PesoPedido'].values[lin])
    Cubagem = str(df['CubagemPedido'].values[lin])
    DataPedido = datai
    QtdVolumes = df['QtdVolumes'].values[lin]

    url = "https://api.ifieldlogistics.com/simulefrete/v1/customer/freight/calculateFreights/"

    querystring = {"originZipcode": CEPOrigem, "destinationZipcode": CEPDestino, "documentClient": CNPJOrigem,
                   "documentReceiver": CNPJDestino, "grossValue": ValorTotal.replace(',','.'), "calculationType": TipoCalculo,
                   "totalWeight": PesoTotal.replace(',','.'), "cubicMeter": Cubagem.replace(',','.'), "orderDate": DataPedido, "haulersLimit": "3",
                   "sortBy": "lowest_price"}

    headers = {
        # 'Authorization': "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiIsImp0aSI6IjU2MDdmZTRiM2RhZjU2ZWZhY2RmZjQzOGJkZTg1Zjk4In0.eyJpc3MiOiJodHRwczpcL1wvYXBpLmlmaWVsZGxvZ2lzdGljcy5jb21cLyIsImF1ZCI6IjE4Ny4zNy4zMS4xOTIiLCJqdGkiOiI1NjA3ZmU0YjNkYWY1NmVmYWNkZmY0MzhiZGU4NWY5OCIsImlhdCI6MTU1NzIzNjgzOCwibmJmIjoxNTU3MjM2ODM4LCJleHAiOjE1NTczMjMyMzgsImNsaWVudE5hbWUiOiJBbmRlcnNvbiBTb3V6YSIsInVzZXJDbGllbnQiOiJhbmRlcnNvbi5zb3V6YSJ9.zc9Zus-pLk8JhMuqnhITpxzWScxrOwWWF42849RxjeM",
        'Authorization': tokenAPI,
        'Cache-Control': "no-cache",
        'Postman-Token': "79fd9f14-7a53-40f1-b010-16e54c4eff5f"
    }

    response2 = requests.request("GET", url, headers=headers, params=querystring)
    j = simplejson.loads(response2.text)

    cont = 0

    erro = str(mid(response2.text,2,12))
    msgerro = str(response.text)

    if erro == 'errorMessage':
        print(contador, Pedido, j)
        lst.append([cont, Pedido, str(j), '', '', '', CNPJDestino, RazaoSocial, '', datai])
        lin = lin + 1
        contador = contador + 1

    else:
        for items in j['pricingHaulerList']:

                try:
                    lst.append([cont, Pedido, str(j["pricingHaulerList"][cont]['tradeHauler']), j["pricingHaulerList"][cont]['documentMainHauler'], j["pricingHaulerList"][cont]['priceFreight'],
                               j["pricingHaulerList"][cont]['deliveryTimeFreight'], CNPJDestino, RazaoSocial, '', datai, Cubagem, QtdVolumes, PesoTotal])
                    cont += 1
                except KeyError:
                    print('ERRO AO SIMULAR O PEDIDO', Pedido, RazaoSocial)
                    cont += 1
        print(contador,'SIMULANDO O FRETE DO PEDIDO', Pedido, RazaoSocial)
        lin = lin + 1
        contador = contador + 1

#criar variavel de moenor preco, nao confiar na ordenacao do sistema

df2 = pd.DataFrame.from_records(lst, columns=colunas)
print('')
print('SIMULAÇÃO FINALIZADA')

# CRIANDO TABELA TEMPORÁRIA SQL
print('')
print('CRIANDO TABELA TEMPORÁRIA DA SIMULAÇÃO')
df2['ValorFrete'].astype(float)
df2.to_sql(
    name='Tabela_SimuleFrete_Temp_v2', # database table name
    con=engine,
    if_exists='replace',
    index=False,
)

# print('EXECUTANDO PROCEDURES  SQL')
# cursor.execute("EXEC Ajustes_SimuleFrete")
# cursor.commit()
# cursor.execute("EXEC RentabilidadeFrete")
# cursor.commit()
# cursor.close()
# print()

print('TABELA CRIADA... PROCESSO FINALIZADO')

