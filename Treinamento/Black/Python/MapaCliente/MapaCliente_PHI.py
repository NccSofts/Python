import pandas as pd
from pandas import ExcelWriter
import pyodbc
import os

desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'OneDrive - Mais Proxima')

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

# # Conexao com o servidor SQL
# engine = sqlalchemy.create_engine(sql_string)
# engine.connect()

# #GERANDO PLANILHAS
arquivo = '\Planilhas\MapaClientes_PHI.xlsx'
desktop = desktop_path + arquivo
writer = ExcelWriter(desktop)
print('Gerando planilha PHI... pode demorar até 5 minutos')
query = "EXEC PowerBIv2.dbo.MapaClientesBI @CanalIMPX = 'PHI'"
print('')
df = pd.read_sql(query, cnxn)
df.to_excel(writer,'Sheet1')
writer.save()
print('Planilha gerada em ' + desktop + ' com sucesso...')
print('-------------------------------------------------------------------------------------------------------------------------')
print('')

