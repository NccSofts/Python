import pyodbc
import sqlalchemy
import pandas as pd
import datetime
from datetime import date


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



print("EXEC PJKTB2..MAPA_CONTAS_RECEBER_IMPX")
cursor.execute("EXEC PJKTB2..MAPA_CONTAS_RECEBER_IMPX")
# while 1:
#     q = status_check_cursor.execute("select RunningStatus from PowerBIv2..Acompanhamento_Rotinas_SQL_Python where Script_Procedure = 'MAPA_CONTAS_RECEBER_IMPX'").fetchone()
#     if q[0] == 0:
#         break
print("Rotina Executada com sucesso")