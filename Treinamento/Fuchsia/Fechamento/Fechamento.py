import pyodbc
import pandas as pd

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


print('Rotina de Fechamento de Custos Logisticos')
print('-----------------------------------------')
print('')

ano = input("Digite o Ano (4 digitos):")
while len(ano)<4:
    print('')
    print('Ano deve ter 4 digitos')
    print('----------------------')
    ano = input("Digite o Ano (4 digitos):")
    print('')

mes = input("Digite o Mes (2 digitos):")
while len(mes)<2:
    print('')
    print('Mês deve ter 4 digitos')
    print('----------------------')
    mes = input("Digite o Mes (2 digitos):")
    print('')

print('')

comando = 'EXEC Datawarehouse..FechamentoLogistica ' + "'" + ano + "'" + ',' + "'" + mes + "'"
#print(comando)

print('Ao executar esta operação todos os dados do respectivo periodo que ja estiverem no banco de dados será apagado...')
confirmacao = input('Deseja mesmo executar o fechamento do período de ' + mes + '/' + ano + ' - (S/N):')
confirmacao = confirmacao.upper()


while confirmacao != 'S' and confirmacao != 'N':
    confirmacao = input('Deseja mesmo executar o fechamento do período de ' + mes + '/' + ano + ' - (S/N):')
    confirmacao = confirmacao.upper()

if confirmacao == 'S':
    print('Executando Fechamento ' + mes + '/' + ano)
    print(comando)
    cursor.execute(comando)
    cursor.commit()

    print('Fechamento Realizado com sucesso...')

if confirmacao == 'N':
    print('Execução de Fechamento Cancelada')

cursor.close()