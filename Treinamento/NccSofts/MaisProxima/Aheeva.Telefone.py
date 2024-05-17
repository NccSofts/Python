import datetime
import sqlalchemy
import sys

import MaisProxima

def sucesso(registros, inicio, fim):
    print('{:%Y-%m-%d %H:%M:%S} [{:d}] Telefone(s) exportado(s) em {:} segundo(s)'
          .format(datetime.datetime.now(), registros, fim - inicio))

def excecao():
    print('{:%Y-%m-%d %H:%M:%S} {:}'
          .format(datetime.datetime.now(), sys.exc_info()[0].__doc__))
    
    email_header = '{:%Y-%m-%d %H:%M:%S} Aheeva.Cliente.py' \
        .format(datetime.datetime.now())
    email_body = '{:}' \
        .format(sys.exc_info()[0].__doc__)
    email_to_email = [ 'siegmar.gieseler@siegmar.com.br' ]
    MaisProxima.email(email_header, email_body, email_to_email) 

connection_sqlserver = None
#transaction_sqlserver = None
connection_mysql = None
transaction_mysql = None

try:
    #exception = 1 / 0

    inicio = datetime.datetime.now()

    registros = 0
    
    # SQL Server

    engine_sqlserver = sqlalchemy.create_engine('mssql+pyodbc://sa:P@ssw0rd@PythoniMPX')
    connection_sqlserver = engine_sqlserver.connect()
    #transaction_sqlserver = connection_sqlserver.begin()

    rows = connection_sqlserver.execute('''
SELECT TOP 100
    COALESCE(CodigoArea, '11') + COALESCE(Telefone, '') Telefone,
    CNPJCPF ChaveCliente
FROM
    Cliente
    INNER JOIN ClienteContato ON
        ClienteContato.IdCliente = Cliente.Id
    INNER JOIN ClienteContato123 ON
        ClienteContato123.IdClienteContato = ClienteContato.Id
        ''')
    
    # MySQL

    engine_mysql = sqlalchemy.create_engine('mssql+pyodbc://sa:P@ssw0rd@PythonTelefone')
    #engine_mysql = sqlalchemy.create_engine('mysql+pyodbc://aheevaccs:aheevaccs@PythonTelefone')
    connection_mysql = engine_mysql.connect()
    transaction_mysql = connection_mysql.begin()

    connection_mysql.execute('TRUNCATE TABLE Telefone')

    sql = sqlalchemy.text('INSERT INTO Telefone VALUES (:Telefone, :ChaveCliente)')
    for row in rows:
        connection_mysql.execute(sql, **row)
        registros += 1
        if (registros % 10 == 0): # 1000 == 0)
            print(registros)

    transaction_mysql.commit()
    #transaction_sqlserver.commit()

    fim = datetime.datetime.now()

    sucesso(registros, inicio, fim)
except:
    if (transaction_mysql is not None):
        transaction_mysql.rollback()
    #if (transaction_sqlserver is not None):
    #    transaction_sqlserver.rollback()

    excecao()
finally:
    if (connection_mysql is not None):
        connection_mysql.close()
    if (connection_sqlserver is not None):    
        connection_sqlserver.close()
