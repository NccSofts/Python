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

query = """\
                UPDATE PowerBiV2..Acompanhamento_Rotinas_SQL_Python 
                    SET RunningStatus = 1
                WHERE Script_Procedure = 'PROC_MAPA_COBRANÇA_IMPX_HISTORICO'
                
                --LIMPA OS REGISTROS JA IMPORTADOS PARA MESMA DATA
                DELETE FROM PowerBIv2..Mapa_Cobrança_IMPX_Historico
                WHERE [DataBase] = CAST(FORMAT(GETDATE(), 'yyyy-MM-dd') AS DATE)
                
                --LIMPA O FILTRO DA DATA ANTERIOR
                UPDATE PowerBIv2..Mapa_Cobrança_IMPX_Historico
                SET FiltroDataAtual = NULL
                WHERE FiltroDataAtual IS NOT NULL
                
                -- INCLUI NOVOS REGISTROS NA BASE
                INSERT INTO PowerBIv2..Mapa_Cobrança_IMPX_Historico
                SELECT * ,CAST(FORMAT(GETDATE(), 'yyyy-MM-dd') AS DATE), 'DataAtual'
                FROM PowerBIv2..Mapa_Cobrança_IMPX 
                
                
                
                
                --CONFIRMA QUE ROTINA EXECUTOU 
                UPDATE PowerBiV2..Acompanhamento_Rotinas_SQL_Python
                
                SET Data = GETDATE(), Horario = FORMAT(GETDATE(), 'HH:mm'), Usuario = (SELECT CURRENT_USER), 
                    [Status] = IIF((SELECT CURRENT_USER) = 'anderson.souza', 'Manual', 'Automatica'),
                    [Processo] = (SELECT DB_NAME() + '..' + OBJECT_NAME(@@PROCID)),
                    Contador = IIF(Contador IS NULL, 0, Contador) + 1,
                    RunningStatus = 0
                
                WHERE Script_Procedure = 'PROC_MAPA_COBRANÇA_IMPX_HISTORICO'
    """




print("EXEC PowerBIV2..PROC_MAPA_COBRANÇA_IMPX_HISTORICO")
cursor.execute('EXEC PowerBIV2..PROC_MAPA_COBRANÇA_IMPX_HISTORICO')
while 1:
    q = status_check_cursor.execute("select RunningStatus from PowerBIv2..Acompanhamento_Rotinas_SQL_Python where Script_Procedure = 'PROC_MAPA_COBRANÇA_IMPX_HISTORICO'").fetchone()
    if q[0] == 0:
        break
print("Rotina Executada com sucesso")


print("EXEC PowerBIV2..PROC_MAPA_COBRANÇA_IMPX_PDD")
cursor.execute('EXEC PJKTB2..PROC_MAPA_COBRANÇA_IMPX_PDD')
while 1:
    q = status_check_cursor.execute("select RunningStatus from PowerBIv2..Acompanhamento_Rotinas_SQL_Python where Script_Procedure = 'PROC_MAPA_COBRANÇA_IMPX_PDD'").fetchone()
    if q[0] == 0:
        break
print("Rotina Executada com sucesso")


