import pandas as pd
import sqlalchemy
import pyodbc
import requests
import simplejson
import os
import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import timedelta, date, time
import json
import urllib.request
import sys
sys.path.append('C:/Python/Black/Python/Database/')
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

def hoje():
    return date.today().strftime('%d/%m/%Y')

today = hoje()


datai = date.today()
dataf = datai - timedelta(15)
dataini = str(datai)
datafinal = str(dataf)


# Definir conexao 0 = PowerBiV2 / 1 = Datawarehouse
cursor = db.sql_conn('PowerBiv2')
engine = db.sql_engine('PowerBiv2')


print('OBTENDO TOKEN DE ACESSO A API')
print('-----------------------------')


print('NCC')


def envia_email_erro(script_python, msg_erro):

    global linha_vazia
    global texto

    user_login = 'bi@maisproxima.com'
    password = '+Proxima2019'
    sender_email = 'Gestão da Informação - Mais Próxima <bi@maisproxima.com>'
    receiver_email = ['ti@maisproxima.com.br']
    titulo_email = 'Erro ao rodar o script ' + str(script_python)

    html = f"""\

                <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Olá,</strong></h1>
                <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;">Houve um erro ao rodar o script {script_python} em {today}.</strong></h1>
                <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;"></strong></h1>
                <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;"></strong>Erro:</h1>
                <h1 style="color: #5e9ca0;"><strong style="color: #000000; font-size: 14px;"></strong>{msg_erro}</h1>
    """

    message = MIMEMultipart("Related")
    message["Subject"] = titulo_email
    message["From"] = sender_email
    message["To"] = ", ".join(receiver_email)
    part = MIMEText(html, "html")
    message.attach(part)

    mailserver = smtplib.SMTP('smtp.gmail.com', 587)
    mailserver.ehlo()
    mailserver.starttls()
    mailserver.login(user_login, password)
    mailserver.sendmail(sender_email, receiver_email, message.as_string())
    mailserver.quit()

    return print('Email do erro enviado')




def token_api():
    try:
        url = "https://api.simulefrete.tec.br/v1/auth"
        # url = "https://sandbox-api.ifieldlogistics.com/v1/auth/"


        # payload = {'clientTokenApi': "TlcoSDRrOXQlSTI4MDIpNjAk.eyJ1c2VybmFtZSI6ImFuZGVyc29uLnNvdXphIiwicGFzc3dvcmQiOiJld2lBZUV3NlRkNVhzUTNWR0pcL3cwQT09In0.SuQZzmPKcO3YlkvVw20h2eHxlIq6F17IKfmnYgH3njg"}
        payload = {'clientTokenApi': "TlcoSDRrOXQlSTI4MDIpNjAk.eyJ1c2VybmFtZSI6ImFwaS5tYWlzcHJveGltYSIsInBhc3N3b3JkIjoiZXdpQWVFdzZUZDVYc1EzVkdKXC93MEE9PSJ9.zqX46gOpETOM_CMJ5t62y5w-c-KhKv6Xt7wzWS2ZUXs"}
        headers = {
            'content-type': "multipart/form-data; boundary=----WebKitFormBoundary7MA4YWxkTrZu0gW",
            'Content-Type': "application/x-www-form-urlencoded",
            'Cache-Control': "no-cache",
            'Postman-Token': "e2061448-9b6c-4a08-b20b-3a5d1b21d41f"
            }

        response = requests.request("POST", url, data=payload, headers=headers)

        numcaract = len(response.text)-14

        tokenAPI = 'Bearer ' + mid(response.text,13,numcaract)
        tokenAPI = tokenAPI.replace('"','')
        print('Token obtido com sucesso')
        print('')
        print('')
        return tokenAPI

    except Exception as error:
        error_string = str(error)
        nome_arquivo = str(os.path.basename(__file__))
        envia_email_erro(nome_arquivo, error_string)
        print('Erro ao tentar obter token de consulta')

token = token_api()

query = 'EXEC PowerBIv2..Pedidos_SimuleFrete'

cursor.execute(query)


lista = """\
                SELECT [Pedido]
                      ,[DataPedido]
                      ,[CNPJ]
                      ,REPLACE([Nome],'.','') Nome
                      ,REPLACE([CEP],'-','') CEP
                      ,[TipoFrete]
                      ,[IdStatusPedido]
                      ,[StatusPedido]
                      ,[QtdVolumes]
                      ,REPLACE([Cubagem], ',', '.') Cubagem
                      ,REPLACE([PesoBruto], ',', '.') PesoBruto
                      ,REPLACE([ValorPedido], ',', '.') ValorPedido
                      ,[CNPJ_Destino]
                      ,REPLACE([Nome_Destino],'.','') Nome_Destino
                      ,REPLACE([CEP_Destino],'-','') CEP_Destino
                      ,[Grupo_Restricao]
                      ,[NF]
                      ,[DataNF]
                      ,[DataConsulta]
                      ,[CNPJ_Transp]
                      ,[Nome_Transp]
                      ,[Preco_Transp]
                      ,[Prazo_Transp]
                      ,[Protheus_Transp]
                      ,[MSG_Erro_TMS]
                FROM [PowerBIv2].[dbo].[Pedidos_Simualacao_Frete_MP]
                --WHERE DataConsulta IS NULL                
                WHERE DataPedido >= '2022-08-01'
                AND Preco_Transp IS NULL
                AND Pedido NOT IN 
                (
                    '426329',
                    '426342',
                    '425267',                
                    '427343',
                    '427410',
                    '427723',
                    '427368',
                    '427734',
                    '427822',
                    '428219',
                    '428338',
                    '428732',
                    '428569',
                    '428975',
                    '428904',
                    '428913',
                    '429026',
                    '429304',
                    '429517',
                    '430201',
                    '430217',            
                    '429454',
                    '430185',
                    --'430186',
                    '430218'

                )               
          """

df = pd.read_sql(lista, engine)


lin = 0

# Cria tabela temporária JSON
colunas = ['Pedido', 'Data_Consulta', 'CNPJ_Tranp', 'Transportadora', 'Frete', 'Prazo_Entrega', 'Grupo_Restricao', 'Ordem']
lst = []

try:
    for row in df.itertuples():

        CNPJOrigem = '11692628000383'
        CEPOrigem = '29162703'
        Pedido = df['Pedido'].values[lin]
        Cliente = df['Nome'].values[lin]
        CNPJDestino = df['CNPJ_Destino'].values[lin]
        CEPDestino = df['CEP_Destino'].values[lin]
        ValorTotal = str(df['ValorPedido'].values[lin])
        PesoTotal = str(df['PesoBruto'].values[lin])
        Cubagem = str(df['Cubagem'].values[lin])
        DataPedido = df['DataPedido'].values[lin]
        QtdVol_int = df['QtdVolumes'].values[lin]
        TipoCalculo = 'total'
        Restricao = df['Grupo_Restricao'].values[lin]
        print('Simulando Frete do Pedido ' + Pedido + ' - ' + Cliente)

        url = "https://api.simulefrete.tec.br/simulefrete/v1/customer/freight/calculateFreights/"

        querystring = {"originZipcode": CEPOrigem, "destinationZipcode": CEPDestino, "documentClient": CNPJOrigem,
                       "documentReceiver": CNPJDestino, "grossValue": ValorTotal, "calculationType": TipoCalculo,
                       "totalWeight": PesoTotal, "cubicMeter": Cubagem, "totalVolumes": QtdVol_int, "orderDate": DataPedido,
                       "haulersLimit": "3", "sortBy": "lowest_price"}

        headers = {
            # 'Authorization': "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiIsImp0aSI6IjU2MDdmZTRiM2RhZjU2ZWZhY2RmZjQzOGJkZTg1Zjk4In0.eyJpc3MiOiJodHRwczpcL1wvYXBpLmlmaWVsZGxvZ2lzdGljcy5jb21cLyIsImF1ZCI6IjE4Ny4zNy4zMS4xOTIiLCJqdGkiOiI1NjA3ZmU0YjNkYWY1NmVmYWNkZmY0MzhiZGU4NWY5OCIsImlhdCI6MTU1NzIzNjgzOCwibmJmIjoxNTU3MjM2ODM4LCJleHAiOjE1NTczMjMyMzgsImNsaWVudE5hbWUiOiJBbmRlcnNvbiBTb3V6YSIsInVzZXJDbGllbnQiOiJhbmRlcnNvbi5zb3V6YSJ9.zc9Zus-pLk8JhMuqnhITpxzWScxrOwWWF42849RxjeM",
            'Authorization': token,
            'Cache-Control': "no-cache",
            'Postman-Token': "79fd9f14-7a53-40f1-b010-16e54c4eff5f"
        }

        response2 = requests.request("GET", url, headers=headers, params=querystring)

        obj = simplejson.loads(response2.text)

        erro = str(mid(response2.text, 2, 12))

        cont = 0

        try:
            reg = len(obj["pricingHaulerList"])
        except:
            reg = 0

        if erro == 'errorMessage':
            print('Erro na consulta do pedido ' + Pedido + '/ Mensagem de erro: ' + str(obj))
            query_erro = """\
                                    
                            UPDATE PowerBiv2..Pedidos_Simualacao_Frete_MP
                                SET MSG_Erro_TMS = ? 
                            WHERE Pedido = ?                                
                            """
            cursor.execute(query_erro, str(obj), Pedido)

        else:

            for items in obj['pricingHaulerList']:

                try:
                    Transportadora = obj["pricingHaulerList"][cont]['tradeHauler']
                    Frete = obj["pricingHaulerList"][cont]['priceFreight']
                    Prazo_Entrega = obj["pricingHaulerList"][cont]['deliveryTimeFreight']
                    CNPJ_Tranp = obj["pricingHaulerList"][cont]['documentMainHauler']
                    #print('---> Transportadora: ' + Transportadora + ' - ' + 'Preço do Frete: '+ str(Frete) + ' - ' + 'Prazo de entrega: ' + str(Prazo_Entrega) + ' dia(s)')

                    lst.append([Pedido, datai, CNPJ_Tranp, Transportadora, Frete, Prazo_Entrega, Restricao, cont])

                    cont += 1



                except KeyError:
                    print('---> Erro ao simular o pedido ' + Pedido + 'do cliente ' + Cliente)
                    cont += 1

            cursor.execute(f'UPDATE PowerBiv2..Pedidos_Simualacao_Frete_MP '
                           f'SET DataConsulta = GETDATE()'
                           f'WHERE Pedido = {Pedido}')

        print('')
        lin = lin + 1

    df = pd.DataFrame.from_records(lst, columns=colunas)
    df.to_sql(
        name='Tabela_Frete_Pedidos_MP', # database table name
        con=engine,
        if_exists='append',
        index=False
    )

    cursor.execute('EXEC PowerBIv2..PROC_Ajustes_SimulacaoFretes')
    
    print('Fim da simulação do frete')

except Exception as error:
    error_string = str(error)
    nome_arquivo = str(os.path.basename(__file__))
    envia_email_erro(nome_arquivo, error_string)