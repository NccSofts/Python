"""
import pyodbc 
import datetime

# Some other example server values are
# server = 'localhost\sqlexpress' # for a named instance
# server = 'myserver,port' # to specify an alternate port
server = '52.67.126.131' 
database = 'DataWareHouse' 
username = 'dw' 
password = 'dw@2020!#' 
# ENCRYPT defaults to yes starting in ODBC Driver 18. It's good to always specify ENCRYPT=yes on the client side to avoid MITM attacks.
#cnxn = pyodbc.connect('DRIVER={ODBC Driver 18 for SQL Server};SERVER='+server+';DATABASE='+database+';ENCRYPT=yes;UID='+username+';PWD='+ password)
#cnxn = pyodbc.connect('DRIVER=SQL Server;SERVER='+server+';DATABASE='+database+';ENCRYPT=yes;UID='+username+';PWD='+ password)
cnxn = pyodbc.connect('DRIVER=SQL Server;SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()


#Sample select query
cursor.execute("SELECT * FROM Industria_Previsao_Data_Entrega") 
row = cursor.fetchone() 
while row: 
    #print(row[0]+row[1]+row[2]+row[3]+row[4]+row[5]+row[6]+row[7])
    #dataNF = dataNF.strftime('%m/%d/%Y,',row[4])
    print(row[0]+' - '+row[1]+' - '+row[2]+' - '+row[3]+' - '+ str(row[4]) +' - '+row[5])
    row = cursor.fetchone()
"""


import pyodbc
import pandas as pd
# Some other example server values are
# server = 'localhost\sqlexpress' # for a named instance
# server = 'myserver,port' # to specify an alternate port
server = '52.67.126.131' 
database = 'DataWareHouse' 
username = 'dw' 
password = 'dw@2020!#' 
cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()
# select 26 rows from SQL table to insert in dataframe.
query = "SELECT * FROM Industria_Previsao_Data_Entrega"
df = pd.read_sql(query, cnxn)
print(df.head(2))