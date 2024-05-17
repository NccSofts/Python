import pandas as pd
import sqlalchemy
import pyodbc
import requests
import simplejson
import sys
sys.path.append('C:/Python/Database/')
import database as db

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
cursor = db.sql_conn('PowerBiv2')
engine = db.sql_engine('PowerBiv2')


query = """\
       
             SELECT 
                Numero,
                '11692628000383' CNPJ_Emitente,
                'MAIS PROXIMA COMERCIAL E DISTRIBUIDORA LTDA' Nome_Emitente,
            CASE
                WHEN TipoFrete = 'C' THEN CNPJ
				WHEN TipoFrete = 'F' AND TranspRedespacho IS NULL THEN CNPJ
                WHEN TipoFrete = 'F' THEN CNPJRedesp
            END AS CNPJ_Destinatario,
            'total' TipoCalculo,
            CASE
                WHEN TipoFrete = 'C' THEN ForCli
				WHEN TipoFrete = 'F' AND TranspRedespacho IS NULL THEN ForCli
				WHEN TipoFrete = 'F' THEN TranspRedespacho			
            END AS Nome_Destinatario,
        
            '29162703' CEP_Emittente,
            'SERRA' Cidade_Emitente,
            'ES' UF_Emitente,
        
            CASE
                WHEN TipoFrete = 'C' THEN CEPCliente
				WHEN TipoFrete = 'F' AND TranspRedespacho IS NULL THEN '07174005'
                WHEN TipoFrete = 'F' AND TipoFretePedido = 'RETIRAES' THEN CEPCliente
				WHEN TipoFrete = 'F' AND TipoFretePedido <>'RETIRAES' THEN CEPRedespacho				
            END AS CEP_Destinatario,
        
            CASE
                WHEN TipoFrete = 'C' THEN R.Cidade
				WHEN TipoFrete = 'F' AND TranspRedespacho IS NULL THEN 'GUARULHOS'
                WHEN TipoFrete = 'F' AND TranspRedespacho IS NOT NULL THEN Cidade_Redespacho			
            END AS Cidade_Destinatario,
        
            CASE
                WHEN TipoFrete = 'C' THEN UFOperacao
				WHEN TipoFrete = 'F' AND TranspRedespacho IS NULL THEN 'SP'
                WHEN TipoFrete = 'F' AND TranspRedespacho IS NOT NULL  THEN UF_Redespacho
            END AS UF_Destinatario,
            
            ROUND(SUM(Peso),2) PesoTotal,
            ROUND(SUM(Volume),2) M3,
            SUM(Qtd_Volumes) QtdVolumes,
            SUM(ReceitaBruta) ValorNF,
            Data DataEmissao,
            FORMAT(GETDATE()+1, 'dd/MM/yyyy') Data_Expedicao,
        
            CASE
                WHEN TipoFrete = 'C' THEN 'CIF'
                WHEN TipoFrete = 'F' THEN 'FOB'
            END AS TipoFrete,
        
            CASE
                WHEN TipoFrete = 'C' THEN IIF(C.Cidade IS NULL, 'Interior', 'Capital')
                WHEN TipoFrete = 'F' THEN 'Capital'
            END AS StatusCidade,
            RegiaoOperacao Regiao,
			TipoFretePedido
        
        FROM [PowerBIv2].[dbo].[Tabela_Rentabilidade_BI] R
        LEFT JOIN PowerBIv2..Tabela_Capitais_Brasil C ON C.Cod_IBGE = R.CidadeX       
        WHERE TipoNF = 'Venda' AND TipoFrete in ('C','F') AND TipoFretePedido <> 'RETIRAES' AND Data >= '2019-10-01' AND VlrFreteProvisao IS NULL --AND LEN(CNPJ) = 14
        --where Numero = '133572.3'
        GROUP BY	Numero, Filial, Operacao, CNPJRedesp,
                    CNPJ, TipoFrete, ForCli, TranspRedespacho,
                    CEPRedespacho, CEPCliente, Cidade_Redespacho, R.Cidade,
                    UFOperacao, UF_Redespacho, Data, C.Cidade, RegiaoOperacao, TipoFretePedido  
        
        """

df = pd.read_sql_query(query, engine)


# tokenAPI = "TlcoSDRrOXQlSTI4MDIpNjAk.eyJ1c2VybmFtZSI6ImFuZGVyc29uLnNvdXphIiwicGFzc3dvcmQiOiJld2lBZUV3NlRkNVhzUTNWR0pcL3cwQT09In0.SuQZzmPKcO3YlkvVw20h2eHxlIq6F17IKfmnYgH3njg"

# SOLICITANDO NOVA KEY DE ACESSO A API

print('OBTENDO TOKEN DE ACESSO A API')
print('-----------------------------')

url = "https://api.simulefrete.tec.br/v1/auth"
# url = "https://sandbox-api.ifieldlogistics.com/v1/auth/"


payload = {'clientTokenApi': "TlcoSDRrOXQlSTI4MDIpNjAk.eyJ1c2VybmFtZSI6ImFwaS5tYWlzcHJveGltYSIsInBhc3N3b3JkIjoiZXdpQWVFdzZUZDVYc1EzVkdKXC93MEE9PSJ9.zqX46gOpETOM_CMJ5t62y5w-c-KhKv6Xt7wzWS2ZUXs"}
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
colunas = ['Order', 'ChavePedido', 'Transportadora', 'CNPJTransp', 'ValorFrete', 'Prazo', 'CNPJCliente', 'RazaoSocial', 'MenorPreco', 'DataConsulta', 'Cubagem','QtdVolumes', 'PesoTotal', 'TipoFrete']
lst = []



# print('IMPORTANDO NOVOS PEDIDOS DE VENDA PARA SIMULACAO')
# # engine.execute("DELETE FROM DataWareHouse.dbo.Tabela_Simula_Frete WHERE StatusPedido <> 'Faturado' AND StatusPedido <> 'Faturar Parcial'")
# cursor.execute("EXEC Pedidos_SimuleFrete")
# cursor.commit()
# print('')
#
# df = pd.read_sql_table('vBaseSimuleFretev3', engine)
# totalreg = str(len(df))
#
# print('SIMULANDO O FRETE DE '+ totalreg + ' PEDIDOS DE VENDA... AGUARDE')
# print('----------------------------------------------------------------')
# print('')

contador = 1
for row in df.itertuples():

    Pedido = df['Numero'].values[lin]
    RazaoSocial = df['Nome_Destinatario'].values[lin]
    CEPOrigem = df['CEP_Emittente'].values[lin]
    CEPDestino = df['CEP_Destinatario'].values[lin]
    CNPJOrigem = df['CNPJ_Emitente'].values[lin]
    CNPJDestino = df['CNPJ_Destinatario'].values[lin]
    ValorTotal = str(df['ValorNF'].values[lin])
    TipoFrete = df['TipoFrete'].values[lin]
    TipoCalculo = df['TipoCalculo'].values[lin]
    PesoTotal = str(df['PesoTotal'].values[lin])
    Cubagem = str(df['M3'].values[lin])
    DataPedido = datai
    QtdVolumes = df['QtdVolumes'].values[lin]
    QtdVol_int = int(df['QtdVolumes'].values[lin])

    url = "https://api.ifieldlogistics.com/simulefrete/v1/customer/freight/calculateFreights/"

    querystring = {"originZipcode": CEPOrigem, "destinationZipcode": CEPDestino, "documentClient": CNPJOrigem,
                   "documentReceiver": CNPJDestino, "grossValue": ValorTotal.replace(',','.'), "calculationType": TipoCalculo,
                   "totalWeight": PesoTotal.replace(',','.'), "cubicMeter": Cubagem.replace(',','.'), "totalVolumes": QtdVol_int, "orderDate": DataPedido,
                   "haulersLimit": "3", "sortBy": "lowest_price"}

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
                               j["pricingHaulerList"][cont]['deliveryTimeFreight'], CNPJDestino, RazaoSocial, '', datai, Cubagem, QtdVolumes, PesoTotal, TipoFrete])

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
#df2['ValorFrete'].astype(float)
df2.to_sql(
    name='Tabela_ProvisaoFrete_Temp',  # database table name
    con=engine,
    if_exists='replace',
    index=False,
)

print('EXECUTANDO PROCEDURES  SQL')
cursor.execute("EXEC Datawarehouse..Ajustes_SimuleFrete_Provisoes")
cursor.close()
print()

print('TABELA CRIADA... PROCESSO FINALIZADO')




