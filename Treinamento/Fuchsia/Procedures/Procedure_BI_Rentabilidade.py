import pyodbc
import sqlalchemy
import pandas as pd
import datetime
from datetime import date

now = datetime.datetime.now()
year = '{:02d}'.format(now.year)
month = '{:02d}'.format(now.month)
day = '{:02d}'.format(now.day)
hour = '{:02d}'.format(now.hour)
minute = '{:02d}'.format(now.minute)
day_month_year = '{}-{}-{}'.format(year, month, day)
primeiro_dia_mes = date.today().replace(day=1)


query_rentabilidade = 'EXEC PowerBIv2..Rentabilidade_PBI ' + "'" + str(primeiro_dia_mes) + "'"
query_contabilidade = 'EXEC MapaContabilidade_Temp ' + year


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


status_check_cursor = pyodbc.connect(
    'DRIVER={ODBC Driver 13 for SQL Server};SERVER=' + server + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password, autocommit=True)
cursor2 = status_check_cursor.cursor()


print('-----------------------------------------')
print('- Executando Procedures BI               ')
print('-----------------------------------------')
print('')


print(query_rentabilidade)
cursor.execute(query_rentabilidade)
while 1:
    q = status_check_cursor.execute("select RunningStatus from PowerBIv2..Acompanhamento_Rotinas_SQL_Python where Script_Procedure = 'Rentabilidade_PBI'").fetchone()
    if q[0] == 0:
        break


print('EXEC DataWareHouse..PerformanceVendedores')
cursor.execute("EXEC DataWareHouse..PerformanceVendedores")
while 1:
    q = status_check_cursor.execute("select RunningStatus from PowerBIv2..Acompanhamento_Rotinas_SQL_Python where Script_Procedure = 'PerformanceVendedores'").fetchone()
    if q[0] == 0:
        break


print('EXEC DataWareHouse..Emails_Vendedores')
cursor.execute("EXEC DataWareHouse..Emails_Vendedores")
while 1:
    q = status_check_cursor.execute("select RunningStatus from PowerBIv2..Acompanhamento_Rotinas_SQL_Python where Script_Procedure = 'Emails_Vendedores'").fetchone()
    if q[0] == 0:
        break


print(query_contabilidade)
cursor.execute(query_contabilidade)
while 1:
    q = status_check_cursor.execute("select RunningStatus from PowerBIv2..Acompanhamento_Rotinas_SQL_Python where Script_Procedure = 'MapaContabilidade_Temp'").fetchone()
    if q[0] == 0:
        break


print("EXEC PowerBiV2..Proc_Razao_Contabil_Temp")
cursor.execute("EXEC PowerBiV2..Proc_Razao_Contabil_Temp")
while 1:
    q = status_check_cursor.execute("select RunningStatus from PowerBIv2..Acompanhamento_Rotinas_SQL_Python where Script_Procedure = 'Proc_Razao_Contabil_Temp'").fetchone()
    if q[0] == 0:
        break


print("EXEC Powerbiv2..DiarizacaoVendas")
cursor.execute("EXEC Powerbiv2..DiarizacaoVendas")
while 1:
    q = status_check_cursor.execute("select RunningStatus from PowerBIv2..Acompanhamento_Rotinas_SQL_Python where Script_Procedure = 'DiarizacaoVendas'").fetchone()
    if q[0] == 0:
        break



print("EXEC PowerBIv2..Analise_Custos_Contabilidade")
cursor.execute("EXEC PowerBIv2..Analise_Custos_Contabilidade")
while 1:
    q = status_check_cursor.execute("select RunningStatus from PowerBIv2..Acompanhamento_Rotinas_SQL_Python where Script_Procedure = 'Analise_Custos_Contabilidade'").fetchone()
    if q[0] == 0:
        break


print("EXEC PowerBiv2..Mapa_Venda_Estoque")
cursor.execute("EXEC PowerBiv2..Mapa_Venda_Estoque")
while 1:
    q = status_check_cursor.execute("select RunningStatus from PowerBIv2..Acompanhamento_Rotinas_SQL_Python where Script_Procedure = 'Mapa_Venda_Estoque'").fetchone()
    if q[0] == 0:
        break


cursor.close()
cursor2.close()
print('')
print('EXECUÇÃO DAS PROCEDURES FINALIZADAS')
