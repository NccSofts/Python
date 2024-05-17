#!pip install -q simplejson
#!pip install -q requests
#!pip install -q pyodbc
#!pip install -q pandas
#!pip install -q numpy
#!pip install py2exe

import pandas as pd
import sqlalchemy
import pyodbc
import requests
import simplejson

#from sqlalchemy import create_engine

# Cria variaveis para filtro automatico das datas
from datetime import timedelta, date

datai = date.today()
dataf = datai - timedelta(15)
dataini = str(datai)
datafinal = str(dataf)

# link = 'https://tms.tecnovia.com.br/v9/dts-transp?periodoEmissao={inicio:' + datafinal + ',termino:' + dataini + '}&format=json'
link = 'https://tms.tecnovia.com.br/v9/dts-transp?periodoEmissao={inicio:' + datafinal + ',termino:' + dataini + '}&incluirValorPreDtsApurado=true'

print('Criando conexao TMS')
print('Buscando CTE´s emitidos entre: '+datafinal+' e '+dataini)
# Conexao com a API Tecnovias
headers = {
    'Authorization': 'Basic QU5ERVJTT04uU09VWkFAQklPU0NBTi5DT00uQlI6VkVMSE8=',
    'Instancia-Id': '282',
    'Content-Type': 'application/json; charset=UTF-8',
    'Accept': 'application/json'}

r = requests.get(link, headers=headers)
# r = requests.get('https://tms.tecnovia.com.br/v9/dts-transp?periodoEmissao={inicio:2019-01-01,termino:2019-03-12}&format=json', headers=headers)

c = r.content
j = simplejson.loads(c)


engine = sqlalchemy.create_engine('mssql+pyodbc://powerbi:powerbi2018@52.67.126.131/DataWareHouse?driver=ODBC+Driver+13+for+SQL+Server')

#df = pd.read_sql_table('Tabela_Frete_TMS_Oficial', engine)

engine.connect()

print('Criando Tabela JSON')
#Cria tabela temporária JSON
colunas = ['Data_CTE', 'Transportadora', 'CTE', 'ValorCTE', 'Destinatario', 'NF', 'NF_Protheus','Numero','Serie','  ValorCTE', 'ValorTMS', 'Divergencia']
lst = []

cont = 0



for item in j:

    try:
        vlr_cte = j[cont]['valorTotalPreDtsDaConciliacao']
    except KeyError:
        vlr_cte = 0

    try:
        vlr_tms = j[cont]['valorTotalPreDtsDaConciliacao']
    except KeyError:
        vlr_tms = 0

    try:
        vlr_div = j[cont]['divergenciaNoFrete']
    except KeyError:
        vlr_div = 0

    # print(item, j[cont]['valorTotalPreDtsDaConciliacao'], j[cont]['valorTotalDtsTranspDaConciliacao'], j[cont]['divergenciaNoFrete'])
    # vlr_tms = j.get('valorTotalPreDtsDaConciliacao', None)
    # vlr_div = j.get('divergenciaNoFrete', None)


    for itemNF in j[cont]['dtTransp']['notasFiscais']:
        lst.append([j[cont]['dtTransp']['emissao'][:10], j[cont]['dtTransp']['emitente']['nome'].replace("'",""), str(j[cont]['dtTransp']['numero']), str(j[cont]['dtTransp']['valorTotal']), j[cont]['dtTransp']['destinatario']['nome'].replace("'",""), itemNF['numeroFormatado'], str(itemNF['numero'])+"."+str(itemNF['serie']), str(itemNF['numero']), str(itemNF['serie']),  str(vlr_cte), str(vlr_tms), str(vlr_div)])

    cont = cont + 1

df = pd.DataFrame.from_records(lst, columns=colunas)
# print('Criando DataFrame TMS')
# print('Quantidade de registros a serem importados: '+ str(df.__len__()))

df.to_sql(
    name='Tabela_Frete_TMS', # database table name
    con=engine,
    if_exists='replace',
    index=False,
)


# Conexao com o servidor SQL
server = '52.67.126.131'
database = 'DataWareHouse'
username = 'powerbi'
password = 'powerbi2018'
cnxn = pyodbc.connect(
   'DRIVER={ODBC Driver 13 for SQL Server};SERVER=' + server + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password)
cursor = cnxn.cursor()

print('Gravando Dados no SQL')

cursor.execute("EXEC ImportarCTE_TMS")
cursor.commit()
cursor.execute("EXEC RentabilidadeFrete")
cursor.commit()
cursor.close()

print('Script Finalizado')


# conn = engine.connect()
# conn.execute("EXEC ImportarCTE_TMS")
# conn.close()