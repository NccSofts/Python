import requests
import simplejson
import pandas as pd
import pyodbc
import sqlalchemy
from datetime import timedelta, date

datai = date.today()
dataf = datai - timedelta(90)
dataini = str(datai)
datafinal = str(dataf)


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
    'DRIVER={ODBC Driver 13 for SQL Server};SERVER=' + server + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password, autocommit=True)
cursor = cnxn.cursor()

# Conexao com o servidor SQL
engine = sqlalchemy.create_engine(sql_string)
engine.connect()


# OBTENDO CHAVE DE ACESSO A API DO I9
print('OBTENDO TOKEN DE ACESSO A API')
print('-----------------------------')
url =  "https://api.simulefrete.tec.br/v1/auth"
# url = "https://sandbox-api.ifieldlogistics.com/v1/auth/"
payload = {
    'clientTokenApi': "TlcoSDRrOXQlSTI4MDIpNjAk.eyJ1c2VybmFtZSI6ImFwaS5tYWlzcHJveGltYSIsInBhc3N3b3JkIjoiZXdpQWVFdzZUZDVYc1EzVkdKXC93MEE9PSJ9.zqX46gOpETOM_CMJ5t62y5w-c-KhKv6Xt7wzWS2ZUXs"}
headers = {
    'content-type': "multipart/form-data; boundary=----WebKitFormBoundary7MA4YWxkTrZu0gW",
    'Content-Type': "application/x-www-form-urlencoded",
    'Cache-Control': "no-cache",
    'Postman-Token': "e2061448-9b6c-4a08-b20b-3a5d1b21d41f"
}
response = requests.request("POST", url, data=payload, headers=headers)
numcaract = len(response.text) - 14
tokenAPI = 'Bearer ' + mid(response.text, 13, numcaract)
tokenAPI = tokenAPI.replace('"', '')
print(tokenAPI)

cnpj_embarcador = '11692628000383'

url = " https://api.simulefrete.tec.br/simulefrete/v1/customer/notaFiscal/deliveryAndPaymentInfo/"

querystring = {"documentClient": cnpj_embarcador, "outDateFrom": datafinal, "outDateTo": dataini}
headers = {
    'Authorization': tokenAPI,
    'Cache-Control': "no-cache",
    'Postman-Token': "79fd9f14-7a53-40f1-b010-16e54c4eff5f"
        }
response2 = requests.request("GET", url, headers=headers, params=querystring)
j = simplejson.loads(response2.text)
erro = str(mid(response2.text, 2, 12))
msgerro = str(response.text)
cont = 0
lin = 0

colunas = ['NF', 'DataExpedicao', 'DataPrevisaoEntrega', 'DataEntrega', 'CodigoOcorrencia','UltimaOcorrencia','DataOcorrencia']
lst = []


colunas2 = ['NF', 'CTE', 'SerieCte', 'NomeTransportadora' ,'TipoCTE', 'DataCTE', 'ValorCTE','ValorCalculoTMS','chave_cte']
lst2 = []

contador = 1
for items in j['notasFiscais']:

    # VERIFICA CRIAR AS VARIÁVEIS
    try:
        previsao_entrega = j["notasFiscais"][cont]['forecastDate']
    except KeyError:
        previsao_entrega = ''

    try:
        data_expedicao = j["notasFiscais"][cont]['outDate']

    except KeyError:
        data_expedicao = ''

    try:
        data_entrega = j["notasFiscais"][cont]['deliveryDate']
    except KeyError:
        data_entrega = ''

    contador = contador + 1


    if erro == 'errorMessage':
        print('Erro')

    else:


        for lastEvent in j["notasFiscais"]:

            try:
                ocorrencia = j["notasFiscais"][cont]["lastEvent"]['nameEvent']
            except:
                ocorrencia = ''


            try:
                codigo_ocorrencia = j["notasFiscais"][cont]["lastEvent"]['codeEvent']
            except:
                codigo_ocorrencia = ''

            try:
                data_ocorrencia = j["notasFiscais"][cont]["lastEvent"]['dateEvent']
            except:
                data_ocorrencia = ''


        numero_nf = str(str(j["notasFiscais"][cont]['number']) + '.' + str(j["notasFiscais"][cont]['serie']))
        data_ocorrencia = left(data_ocorrencia,10)
        data_expedicao = left(data_expedicao, 10)
        lst.append([numero_nf, data_expedicao, previsao_entrega, data_entrega, codigo_ocorrencia ,ocorrencia , data_ocorrencia])
        cont = cont + 1
        print('Dados de tracking encontrados para a nf', numero_nf)

        i = 0


        try:

            for each in j["notasFiscais"][cont]["ctes"]:
                numero_nf = str(str(j["notasFiscais"][cont]['number']) + '.' + str(j["notasFiscais"][cont]['serie']))
                cte = str(j["notasFiscais"][cont]["ctes"][i]['number'])
                nome_tranportadora = str(j["notasFiscais"][cont]["ctes"][i]['tradeHauler'])

                try:
                    serie_cte = str(j["notasFiscais"][cont]["ctes"][i]['serie'])
                except:
                    serie_cte = None


                try:
                    chave_cte = str(j["notasFiscais"][cont]["ctes"][i]['key'])
                except:
                    chave_cte = None

                tipo_cte = j["notasFiscais"][cont]["ctes"][i]['type']
                data_cte = j["notasFiscais"][cont]["ctes"][i]['issueDate']
                valor_cte = j["notasFiscais"][cont]["ctes"][i]['chargedGrossFreightNotaFiscal']
                valor_calc_tms = j["notasFiscais"][cont]["ctes"][i]['calculatedGrossFreightNotaFiscal']


                lst2.append([numero_nf, cte, serie_cte, nome_tranportadora, tipo_cte, data_cte, valor_cte, valor_calc_tms, chave_cte])
                i +=1
                print('CTE número ', cte, 'encontrado para a nf', numero_nf)

        except:
            print('CTE não encontrado para a NF', numero_nf)



df = pd.DataFrame.from_records(lst, columns=colunas)
df2 = pd.DataFrame.from_records(lst2, columns=colunas2)


print('')
print('CRIANDO TABELA TEMPORÁRIA DE TRACKING DE NFS')
df.to_sql(
    name='Tabela_SimuleFrete_Tracking_TEMP', # database table name
    con=engine,
    if_exists='replace',
    index=False,
)

print('')
print('CRIANDO TABELA TEMPORÁRIA DE CTES')
df2.to_sql(
    name='Tabela_SimuleFrete_CTES_TEMP', # database table name
    con=engine,
    if_exists='replace',
    index=False,
)

print('')
print('INSERINDO NOVAS NFS NA BASE Tabela_SimuleFrete_ CTES E TRACKING E ATUALIZANDO DADOS')
query = """\
        --INSERINDO NOVAS NFS NA BASE DE TRACKING OFICIAL
        INSERT INTO PowerBIv2..Tabela_SimuleFrete_Tracking (NF, DataNF, Ano, Mes, TipoFrete) 
        SELECT 
            T.Numero,
            T.Data, 
            T.Ano, 
            T.Mes,
            T.TipoFrete
        
        FROM PowerBIv2..Tabela_Rentabilidade_BI T
        LEFT JOIN PowerBIv2..Tabela_SimuleFrete_Tracking R ON R.NF = T.Numero AND R. DataNF = T. Data
        WHERE T.Data >= '2019-09-01' AND T.TipoNF = 'Venda' AND T.TipoFrete in ('C', 'F') AND R.NF IS NULL
        GROUP BY T.Numero, T.Data, T.Ano, T.Mes, T.TipoFrete
        
        
        --ATUALIZANDO OS DADOS DA TABELA OFICIAL
        UPDATE PowerBIv2..Tabela_SimuleFrete_Tracking
        SET DataEntrega = IIF(T.DataEntrega = '', NULL, T.DataEntrega),
            DataExpedicao = IIF(T.DataExpedicao = '', NULL, T.DataExpedicao),
            DataPrevisaoEntrega = IIF(T.DataPrevisaoEntrega = '', NULL, T.DataPrevisaoEntrega),
            CodigoOcorrencia = T.CodigoOcorrencia,
            UltimaOcorrencia = T.UltimaOcorrencia,
            DataOcorrencia = IIF(T.DataOcorrencia = '', NULL, T.DataOcorrencia)
        
        FROM PowerBIv2..Tabela_SimuleFrete_Tracking O
        INNER JOIN PowerBIv2..Tabela_SimuleFrete_Tracking_TEMP T ON T.NF = O.NF
        
        
        --INSERINDO NOVOS CTES NA BASE OFICIAL
        INSERT INTO PowerBIv2..Tabela_SimuleFrete_CTES (NF, CTE,  SerieCTE, NomeTransportadora, TipoCTE, DataCTE, ValorCTE, ValorCalculoTMS, Chave_CTE)
        SELECT 
            O.NF, O.CTE,  O.SerieCTE, O.NomeTransportadora, O.TipoCTE, O.DataCTE, O.ValorCTE, O.ValorCalculoTMS, O.chave_cte
        
        FROM PowerBIv2..Tabela_SimuleFrete_CTES_TEMP  O
        LEFT JOIN PowerBIv2..Tabela_SimuleFrete_CTES T ON T.NF = O.NF AND T.CTE = O.CTE
        WHERE T.NF IS NULL
        
        
        -- ATUALIZANDO DAODS DE ENTREGA NFS
        UPDATE PowerBIv2..Tabela_Rentabilidade_BI
        SET DataExpedicao = IIF(R.DataExpedicao IS NULL, T.DataExpedicao, R.DataExpedicao),
            DataPrevisaoEntrega = IIF(R.DataPrevisaoEntrega IS NULL, T.DataPrevisaoEntrega, R.DataPrevisaoEntrega),
            DataEntrega = IIF(R.DataEntrega IS NULL, T.DataEntrega, R.DataEntrega)
        
        FROM PowerBIv2..Tabela_Rentabilidade_BI R
        INNER JOIN PowerBIv2..Tabela_SimuleFrete_Tracking T ON T.NF = R.Numero
        WHERE IIF(R.DataExpedicao IS NULL, 'A',IIF(R.DataPrevisaoEntrega IS NULL, 'B', IIF(R.DataEntrega IS NULL, 'C','N'))) IN ('A', 'B', 'C')


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

        --VERIFICA DATA CTE NO PROTHEUS
        UPDATE [PowerBIv2].[dbo].[Tabela_SimuleFrete_CTES]

            SET Data_Classificacao_Protheus = Data

        FROM [PowerBIv2].[dbo].[Tabela_SimuleFrete_CTES] T
        INNER JOIN DataWareHouse..MapaFiscal M ON RTRIM(M.ChaveNFE) = LEFT(T.Chave_CTE,45)
        WHERE Data_Classificacao_Protheus IS NULL


        --VERIFICA DATA CTE NO PROTHEUS
        UPDATE [PowerBIv2].[dbo].[Tabela_SimuleFrete_CTES]

            SET Data_Classificacao_Protheus = Data

        FROM [PowerBIv2].[dbo].[Tabela_SimuleFrete_CTES] T
        INNER JOIN DataWareHouse..MapaFiscal M ON RTRIM(M.ChaveNFE) = LEFT(T.Chave_CTE,44)
        WHERE Data_Classificacao_Protheus IS NULL
        
        
        --VERIFICA DATA CTE NO PROTHEUS
        UPDATE [PowerBIv2].[dbo].[Tabela_SimuleFrete_CTES]

            SET Data_Classificacao_Protheus = Data

        FROM [PowerBIv2].[dbo].[Tabela_SimuleFrete_CTES] T
        INNER JOIN DataWareHouse..MapaFiscal M ON RTRIM(M.ChaveNFE) = RTRIM(T.Chave_CTE)
        WHERE Data_Classificacao_Protheus IS NULL


        --PRINT('RATEIA FRETE REAL NA NF')
        UPDATE PowerBIv2..Tabela_Rentabilidade_BI
        SET CTE = F.CTE, Transportadora = F.NomeTransportadora, DataCTEProtheus = F.DataCTE, VlrFreteReal =  -F.ValorCTE * IndiceRateioCubagem
        FROM PowerBIv2..Tabela_Rentabilidade_BI R
        INNER JOIN PowerBIv2..Tabela_SimuleFrete_CTES F ON	F.NF = R.Numero


        -- INCLUI NOVOS REGISTROS NA TABELA DE ACOMPANHAMENTO DE FRETE ANO X ANO (BONUS LOGISTICA)
        INSERT INTO [PowerBIv2].[dbo].[AcompanhamentoFrete] (	   [Numero]
                                                                  ,[Cliente]
                                                                  ,[Chave]
                                                                  ,[Regiao]
                                                                  ,[Mes]
                                                                  ,[Cidade]
                                                                  ,[UF]
                                                                  ,[TipoFrete]
                                                                  ,[Grupo]
                                                                  ,[ReceitaBruta]
                                                                  ,[FreteAnoAtual]
                                                                  ,[FreteColeta]
                                                                  ,[Coleta]
                                                                  ,[Cubagem]
                                                                  ,[Peso]
                                                                  ,[FreteAnoAnterior]
                                                                  ,[AnoAtual]
                                                                  ,[TipoCalculo])
        SELECT

               V.[Numero]
              ,V.[Cliente]
              ,V.[Chave]
              ,V.[Regiao]
              ,V.[Mes]
              ,V.[Cidade]
              ,V.[UF]
              ,V.[TipoFrete]
              ,V.[Grupo]
              ,V.[ReceitaBruta]
              ,V.[FreteAnoAtual]
              ,V.[FreteColeta]
              ,V.[Coleta]
              ,V.[Cubagem]
              ,V.[Peso]
              ,V.[FreteAnoAnterior]
              ,V.[AnoAtual]
              ,V.[TipoCalculo]

        FROM [PowerBIv2].[dbo].[v_AcompanhamentoFreteBD] V
        LEFT JOIN [PowerBIv2].[dbo].[AcompanhamentoFrete] T ON T.Numero = V.Numero
        WHERE T.Numero IS NULL

        --CONFIRMA QUE ROTINA EXECUTOU 
        UPDATE PowerBiV2..Acompanhamento_Rotinas_SQL_Python
        
        SET Data = GETDATE(), Horario = FORMAT(GETDATE(), 'HH:mm'), Usuario = (SELECT CURRENT_USER), 
            [Status] = IIF((SELECT CURRENT_USER) = 'anderson.souza', 'Manual', 'Automatica'),
            Processo = (SELECT OBJECT_SCHEMA_NAME(@@PROCID) + '.' + OBJECT_NAME(@@PROCID)),
            Contador = Contador + 1
        WHERE Script_Procedure = 'Importar_CTE_I9_Fretes'

        
    """
cursor.execute(query)
cursor.close()
print('FINAL DAS IMPORTAÇÕES')