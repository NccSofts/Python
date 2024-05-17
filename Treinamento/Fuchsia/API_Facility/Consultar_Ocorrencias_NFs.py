import requests
import simplejson
import pandas as pd
import pyodbc
import sqlalchemy

def left(s, amount):
    return s[:amount]

def right(s, amount):
    return s[-amount:]

def mid(s, offset, amount):
    return s[offset:offset+amount]

# Definir conexao 0 = PowerBiV2 / 1 = Datawarehouse
opcao = 0

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


# OBTENDO CHAVE DE ACESSO A API DO I9
print('OBTENDO TOKEN DE ACESSO A API')
print('-----------------------------')
url = "https://api.ifieldlogistics.com/v1/auth"
# url = "https://sandbox-api.ifieldlogistics.com/v1/auth/"
payload = {
    'clientTokenApi': "TlcoSDRrOXQlSTI4MDIpNjAk.eyJ1c2VybmFtZSI6ImFuZGVyc29uLnNvdXphIiwicGFzc3dvcmQiOiJld2lBZUV3NlRkNVhzUTNWR0pcL3cwQT09In0.SuQZzmPKcO3YlkvVw20h2eHxlIq6F17IKfmnYgH3njg"}
headers = {
    'Content-type': "multipart/form-data; boundary=----WebKitFormBoundary7MA4YWxkTrZu0gW",
    'Content-Type': "application/x-www-form-urlencoded",
    'Cache-Control': "no-cache",
    'Postman-Token': "e2061448-9b6c-4a08-b20b-3a5d1b21d41f"
}
response = requests.request("POST", url, data=payload, headers=headers)
numcaract = len(response.text) - 14
tokenAPI = 'Bearer ' + mid(response.text, 13, numcaract)
tokenAPI = tokenAPI.replace('"', '')
print(tokenAPI)





query_nfs = """\
	    
		SELECT
            CAST(LEFT(Numero, 6) AS INT) NF,
            R.Numero,
            Data DataNF,
            '11692628000383' CNPJ_Embarcador,
            R.TipoFrete,
            ForCli Cliente,
            FORMAT(GETDATE(), 'yyyy-MM-dd') Hoje
        FROM 
            PowerBIv2..Tabela_Rentabilidade_BI R
        
        LEFT JOIN 
            [PowerBIv2].[dbo].[Tabela_SimuleFrete_Tracking] T ON T.NF = R.Numero
        WHERE 
            TipoNF = 'Venda' AND R.TipoFrete in ('C', 'F') AND Data >= '2019-09-01'
            --AND T.UltimaOcorrencia IS NULL		
            AND IIF(R.DataExpedicao IS NULL, 'A',IIF(R.DataPrevisaoEntrega IS NULL, 'B', IIF(R.DataEntrega IS NULL, 'C','N'))) IN ('A', 'B','C')
        GROUP BY R.Numero, CAST(LEFT(R.Numero, 6) AS INT), R.TipoFrete, ForCli,  Data 
        ORDER BY NF Desc
        
            """
df = pd.read_sql_query(query_nfs, engine)


# Cria tabela temporária JSON
colunas = ['NF', 'DataNF',  'Cliente' , 'TipoFrete', 'CodigoOcorrencia', 'Ocorrencia', 'DataOcorrencia', 'DataConsulta']
lst = []



contador = 1
lin = 0

for row in df.itertuples():

    nf = df['NF'].values[lin]
    cnpj_embarcador = df['CNPJ_Embarcador'].values[lin]
    numero_nf = df['Numero'].values[lin]
    data_nf = df['DataNF'].values[lin]
    cliente = df['Cliente'].values[lin]
    tipofrete = df['TipoFrete'].values[lin]
    data_consulta = df['Hoje'].values[lin]


    url = "https://api.ifieldlogistics.com/simulefrete/v1/customer/notaFiscal/listEvents/"

    querystring = {"notaFiscal": nf, "documentClient": cnpj_embarcador, "showOnlyTracking}": 0}
    # querystring = {"notaFiscal": '129787.3', "documentClient": '11692628000383', "showOnlyTracking}": 0}

    headers = {
        'Authorization': tokenAPI,
        'Cache-Control': "no-cache",
        'Postman-Token': "79fd9f14-7a53-40f1-b010-16e54c4eff5f"
            }

    response2 = requests.request("GET", url, headers=headers, params=querystring)
    j = simplejson.loads(response2.text)

    erro = str(mid(response2.text,2,12))
    msgerro = str(response.text)
    cont = 0

    if erro == 'errorMessage':
        print(contador,' NF ', numero_nf,' não encontrada')
        contador = contador + 1


    else:
        for items in j['events']:
            try:
                data_evento = left(j["events"][cont]['dateEvent'],10)
                lst.append([numero_nf, data_nf, cliente, tipofrete, str(j["events"][cont]['codeEvent']), str(j["events"][cont]['nameEvent']), data_evento, data_consulta])
                print(contador, numero_nf, str(j["events"][cont]['nameEvent']), data_evento)
                cont += 1
                contador = contador + 1
            except KeyError:
                print('NF não encontrada')
    # lst = lst.sort()
    #     contador = contador + 1
    lin = lin + 1


df2 = pd.DataFrame.from_records(lst, columns=colunas)

df2.to_sql(
    name='Ocorrencias_Transporte_NFs_TEMP', # database table name
    con=engine,
    if_exists='replace',
    index=False,
)

# ATUALIZANDO TABELAS SQL
print('INSERINDO NOVOS DADOS NA TABELA OFICIAL')
query = """\

            INSERT INTO PowerBIv2..Ocorrencias_Transporte_NFs (NF, DataNF, Cliente, TipoFrete, CodigoOcorrencia, Ocorrencia, DataOcorrencia, DataConsulta)

            SELECT
                F.NF,
                F.DataNF,
                F.Cliente,
                F.TipoFrete,
                F.CodigoOcorrencia,
                UPPER(F.Ocorrencia),
                F.DataOcorrencia,
                F.DataConsulta

            FROM PowerBIv2..Ocorrencias_Transporte_NFs_TEMP F
            LEFT JOIN PowerBIv2..Ocorrencias_Transporte_NFs T ON T.NF = F.NF AND T.CodigoOcorrencia = F.CodigoOcorrencia
            WHERE T.CodigoOcorrencia IS NULL

            """
cursor.execute(query)
cursor.commit()


# ATUALIZANDO TABELAS
print('ATUALIZANDO OS DADOS DE ENTREGA NA RENTABILIDADE')
query = """\

            --  ATUALIZA RENTABILIDADE
            UPDATE PowerBiv2..Tabela_Rentabilidade_BI

            SET DataExpedicao = E.DataOcorrencia

            FROM PowerBiv2..Tabela_Rentabilidade_BI R
            INNER JOIN(
                        SELECT [NF]
                              ,[DataNF]
                              ,[Cliente]
                              ,[TipoFrete]
                              ,[CodigoOcorrencia]
                              ,[Ocorrencia]
                              ,[DataOcorrencia]
                              ,[DataConsulta]
                          FROM PowerBiv2..Ocorrencias_Transporte_NFs_TEMP
                          WHERE CodigoOcorrencia = '00'
                          ) E ON E.NF = R.Numero

            WHERE DataExpedicao IS NULL


            UPDATE PowerBiv2..Tabela_Rentabilidade_BI

            SET DataEntrega = E.DataOcorrencia

            FROM PowerBiv2..Tabela_Rentabilidade_BI R
            INNER JOIN(
                        SELECT [NF]
                              ,[DataNF]
                              ,[Cliente]
                              ,[TipoFrete]
                              ,[CodigoOcorrencia]
                              ,[Ocorrencia]
                              ,[DataOcorrencia]
                              ,[DataConsulta]
                          FROM PowerBiv2..Ocorrencias_Transporte_NFs_TEMP
                          WHERE CodigoOcorrencia = '01'
                          ) E ON E.NF = R.Numero

            WHERE DataEntrega IS NULL

            --ATUALIZAR STATUS DA ENTREGA
            UPDATE PowerBIv2.dbo.Tabela_Rentabilidade_BI
            SET StatusEntrega = E.StatusEntrega, StatusPrazoEntrega = E.StatusPrevisaoEntrega

            FROM PowerBIv2.dbo.Tabela_Rentabilidade_BI R
            INNER JOIN(
            SELECT
                    Numero,

                    CASE
                        WHEN DataEntrega IS NOT NULL THEN 'Entregue'
                        WHEN DataExpedicao IS NULL AND DataPrevisaoEntrega IS NOT NULL THEN 'Pendente de Entrega'
                        WHEN CTE <> '' AND DataEntrega IS NULL THEN 'Pendente de Entrega'
                        WHEN DataExpedicao IS NULL AND DataEntrega IS NULL AND DataPrevisaoEntrega IS NULL THEN 'Aguardando Expedição ou Retira'
                        WHEN DataExpedicao = '' AND DataEntrega = '' AND DataPrevisaoEntrega = '' THEN 'Aguardando Expedição ou Retira'
                        WHEN DataExpedicao IS NULL THEN 'Aguardando Expedição ou Retira'
                        WHEN DataExpedicao IS NOT NULL AND DataEntrega IS NULL THEN 'Pendente de Entrega'
                    END AS StatusEntrega,

                    CASE
                        WHEN DataEntrega > DataPrevisaoEntrega THEN 'Fora do Prazo'
                        WHEN DataEntrega <= DataPrevisaoEntrega THEN 'No Prazo'
                        WHEN DataEntrega IS NULL AND DataPrevisaoEntrega >= CAST(GETDATE() AS DATE) THEN 'No Prazo'
                        WHEN DataEntrega IS NULL AND DataPrevisaoEntrega <  CAST(GETDATE() AS DATE) THEN 'Fora do Prazo'
                        WHEN DataPrevisaoEntrega IS NULL THEN 'Sem Previsão Informada'
                    END AS StatusPrevisaoEntrega

            FROM PowerBIv2.dbo.Tabela_Rentabilidade_BI
            WHERE TipoNF = 'Venda'
            GROUP BY Numero, DataExpedicao, DataEntrega, DataPrevisaoEntrega, CTE
            ) E ON E.Numero = R.Numero


            UPDATE PowerBIv2.dbo.Tabela_Rentabilidade_BI
            SET StatusEntrega = 'Pendente de Entrega'
            FROM PowerBIv2.dbo.Tabela_Rentabilidade_BI
            WHERE StatusEntrega = 'Aguardando Expedição ou Retira' AND CTE IS NOT NULL


            --CONFIRMA QUE ROTINA EXECUTOU
            UPDATE PowerBiV2..Acompanhamento_Rotinas_SQL_Python

            SET Data = GETDATE(), Horario = FORMAT(GETDATE(), 'HH:mm'), Usuario = (SELECT CURRENT_USER),
                [Status] = IIF((SELECT CURRENT_USER) = 'anderson.souza', 'Manual', 'Automatica'),
            Processo = (SELECT OBJECT_SCHEMA_NAME(@@PROCID) + '.' + OBJECT_NAME(@@PROCID)),
            Contador = Contador + 1
            WHERE Script_Procedure = 'Consultar_Ocorrencias_NFs'


            """
cursor.execute(query)
cursor.commit()
cursor.close()