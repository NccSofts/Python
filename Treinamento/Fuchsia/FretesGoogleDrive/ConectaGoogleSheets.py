#!pip install gspread oauth2client


#IMPORTACAO PLANILHA FRETE COTACAO
import gspread
import pyodbc
import sqlalchemy
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd


from datetime import timedelta, date

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



scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secrets.json', scope)
client = gspread.authorize(creds)

sheet = client.open('Fretes_Cotaçao').sheet1


Planilha1 = sheet.get_all_records()

colunas1 = ['NFzzz', 'Clientezzz', 'Ufzzz', 'iPDzzz', 'Total', 'Frete Cotado', 'Data','Tipo']
df = pd.DataFrame.from_records(Planilha1, columns=colunas1)
df = df.astype(str)

print('IMPORTACAO PLANILHA FRETE COTACAO')
print('')
print('    Quantidade de registros a serem importados de Frete_Cotacao: '+ str(df.__len__()))
print('    Gravando na tabela SQL Tabela_Frete_Cotacao_Temp')

df.to_sql(
    name='Tabela_Frete_Cotacao_Temp', # database table name
    con=engine,
    if_exists='replace',
    index=False,
)

hoje = date.today()
hojef = hoje.strftime('%d/%m/%Y')
texto = 'Importado para SQL em '+str(hojef)
lin = 2
cell = 9

#Grava na planilha que o registro foi importado para o SQL
for items in Planilha1:
    sheet.update_cell(lin, cell, texto)
    lin = lin + 1


print('    Exportando Fretes Cotados para RentabilidadeAjuaste')
cursor.execute("EXEC FreteCotado")
cursor.commit()



#IMPORTACAO PLANILHA FRETE COLETA

print('')
print('')
print('IMPORTACAO PLANILHA FRETE COLETA')
print('')
sheet = client.open('Frete_Coleta').sheet1
Planilha1 = sheet.get_all_records()

colunas1 = ['NumeroNF', 'Fornecedor', 'ValorNF', 'Transportadora', 'DataColeta', 'CTE', 'ValorFrete', 'TipoVeículo', 'CubagemCarga', 'PesoCarga','Tipo']
df = pd.DataFrame.from_records(Planilha1, columns=colunas1)
df = df.astype(str)


print('    Quantidade de registros a serem importados de Frete_Coleta: '+ str(df.__len__()))
print('    Gravando na tabela SQL Tabela_Frete_Coleta_Temp')

df.to_sql(
    name='Tabela_Frete_Coleta_Temp', # database table name
    con=engine,
    if_exists='replace',
    index=False,
)


print('    Exportando Fretes Coleta para Tabela_Coletas_MP2')
cursor.execute("EXEC FreteColeta")
cursor.commit()


hoje = date.today()
hojef = hoje.strftime('%d/%m/%Y')
texto = 'Importado para SQL em '+str(hojef)
lin = 2
cell = 12

#Grava na planilha que o registro foi importado para o SQL
for items in Planilha1:
    sheet.update_cell(lin, cell, texto)
    lin = lin + 1


cursor.close()

print('')
print('')
print('Fim do Script')
