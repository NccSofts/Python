import datetime
import sqlalchemy
import sys

import MaisProxima

# Aheeva MySQL 5.1.73

def sucesso(registros, inicio, fim):
    print("{:%Y-%m-%d %H:%M:%S} [{:d}] Cliente(s) exportado(s) em {:} segundo(s)"
          .format(datetime.datetime.now(), registros, fim - inicio))

def excecao():
    print("{:%Y-%m-%d %H:%M:%S} {:}"
          .format(datetime.datetime.now(), sys.exc_info()[0].__doc__))
    
    email_header = "{:%Y-%m-%d %H:%M:%S} Aheeva.Cliente.py" \
        .format(datetime.datetime.now())
    email_body = "{:}" \
        .format(sys.exc_info()[0].__doc__)
    email_to_email = [ "siegmar.gieseler@siegmar.com.br" ]

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

    print("create_engine 1")
    engine_sqlserver = sqlalchemy.create_engine("mssql+pyodbc://impx:9UfRl8hlvlSwUqlduY*z@PythoniMPX")
    print("connect 1")
    connection_sqlserver = engine_sqlserver.connect()
    #transaction_sqlserver = connection_sqlserver.begin()

    if False:
        rows = connection_sqlserver.execute("""
SELECT
    CNPJCPF Chave,
    Cliente.Nome Nome,
    Municipio.Nome Cidade,
    UF.Sigla UF,
    "Vendedor" Vendedor
FROM
    Cliente
    INNER JOIN Municipio ONS
        Municipio.Id = Cliente.IdMunicipio
    INNER JOIN UF ON
        UF.Id = Municipio.IdUF
""")
    
    # MySQL

    engine_mysql = sqlalchemy.create_engine("mysql+mysqldb://aheevaccs:aheevaccs@179.191.102.82/telefone")
    #engine_mysql = sqlalchemy.create_engine("mysql+mysqldb://aheevaccs:aheevaccs@179.191.114.170/telefone")
    connection_mysql = engine_mysql.connect()
    transaction_mysql = connection_mysql.begin()

    connection_mysql.execute("TRUNCATE TABLE Cliente")

    sql = sqlalchemy.text("INSERT INTO Cliente VALUES (:Chave, :Nome, :Cidade, :UF, :Vendedor)")
    for row in rows:
        connection_mysql.execute(sql, row)
        registros += 1
        print(registros)
        if (registros % 100 == 0):
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
