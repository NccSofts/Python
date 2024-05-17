# from Database.database import sql_conn, sql_engine
# cursor = sql_conn('PowerBiv2')
# engine = sql_engine('PowerBiv2')
#
# cursor.execute("EXEC DataWareHouse.dbo.IndustriaPedidoNFe")
# print('Procedure Finalizada')

import pandas as pd
import numpy as np


df = pd.DataFrame(np.array([[1, 2], [1, 5], [0, 8], [0, 10], [1, 13], [1,11],]), columns=['a', 'b'])
df['c'] = np.where(df['a']== 1, df['b'] , 0)
dd=[]

val_anterior = 0
val_atual = 0
indice_anterior = 0
lin = 0

for items in df.itertuples():
#    if val_anterior == 0.1:
#        val_anterior = df['c'].values[lin]
#        val_atual = val_anterior
#        indice_anterior = df['a'].values[lin]

    if indice_anterior == 0 and df['a'].values[lin] == 0:
        val_atual = 0
        indice_anterior = df['a'].values[lin]

    if indice_anterior == 0 and df['a'].values[lin] == 1:
        val_anterior = df['c'].values[lin]
        val_atual = val_anterior
        indice_anterior = df['a'].values[lin]

    if indice_anterior == 1 and df['a'].values[lin] == 1:
        val_atual = val_anterior
        indice_anterior = df['a'].values[lin]

    if indice_anterior == 1 and df['a'].values[lin] == 0:
        val_atual = 0
        indice_anterior = df['a'].values[lin]

    print(df['a'].values[lin], df['b'].values[lin], indice_anterior, val_atual)
    dd.append(val_atual)

    lin = lin + 1

print(dd)
df['d'] = dd

print(df)

# print(df)