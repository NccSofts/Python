import datetime
import os
import pandas as pd
import time
import pyodbc
import sqlalchemy
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import  Keys
from datetime import timedelta, date
import time
import sys
sys.path.append('C:/Python/Database/')
import database as db


def check_exists_by_xpath(xpath):
    try:
        browser.find_element_by_xpath(xpath)
    except NoSuchElementException:
        return False
    return True

def check_exists_by_name(name):
    try:
        browser.find_element_by_name(name)
    except NoSuchElementException:
        return False
    return True

def check_exists_by_id(name):
    try:
        browser.find_element_by_id(name)
    except NoSuchElementException:
        return False
    return True


def frame_switch(name):
  browser.switch_to.frame(browser.find_element_by_name(name))

def substring(s, start, end):
    return s[start:end]

# Data de hoje
hoje = db.hoje()
datai = date.today()
data_query = date.today().strftime('%Y%m%d')
now = datetime.datetime.now()
data_hora = time.strftime('%d/%m/%Y - %H:%M')


# Definir conexao 0 = PowerBiV2 / 1 = Datawarehouse
cursor = db.sql_conn('PowerBiv2')
engine = db.sql_engine('PowerBiv2')

lista = """\
            SELECT 
                NF,
				SERIE_NF,
                Pedido,
                [Data Faturamento],
                [Previsao Entrega]
            FROM PowerBIv2..Acompanhamento_Industrias_Base_v2
            WHERE Industria = 'Lenoxx' AND
                    --Status1 IN ('Entregue') AND
                    Status1 IN ('Pendente Entrega', 'Pendente Entrega em atraso') AND
                    LEN(NF) >= 6
            GROUP BY NF, Pedido, [Data Faturamento], [Previsao Entrega], SERIE_NF
            ORDER BY [Previsao Entrega]
          """
df = pd.read_sql(lista, engine)

lin = 0

delete = """\
                DELETE FROM [PowerBIv2].[dbo].[Tabela_Consulta_Pedidos_Lenoxx]
                WHERE [StatusEntrega] IN ('Entrega não registrada no sistema')
          """
cursor.execute(delete)





#Conexao HOST Selenium
chrome_options = webdriver.ChromeOptions();
chrome_options.add_experimental_option("excludeSwitches", ['enable-logging']);

browser = webdriver.Chrome(executable_path=r'C:/Python/Selenium/chromedriver.exe', options=chrome_options)  # Optional argument, if not specified will search path
browser.maximize_window()
action = ActionChains(browser)

browser.get('https://portal.lenoxx.com.br/portal_lenoxx/autenticar_usuario.php')
# usuario = browser.find_element_by_id('usuario')
# senha = browser.find_element_by_id('senha')

usuario = browser.find_element_by_xpath('//*[@id="_easyui_textbox_input1"]')
senha = browser.find_element_by_xpath('//*[@id="_easyui_textbox_input2"]')
botao_entrar = browser.find_element_by_xpath('//*[@id="form_login"]/table/tbody/tr/td[2]/table/tbody/tr[4]/td/a/span/span[1]')

usuario.send_keys('igor.chaves@maisproxima.com.br')
senha.send_keys('191297')
botao_entrar.click()

cursor.execute('TRUNCATE TABLE PowerBiv2..Industrias_Data_Entrega_TEMP ')

lin = 0
for items in df.itertuples():
    # ABRE BROWSER
    nf = df['NF'].values[lin]
    serie_nf = df['SERIE_NF'].values[lin]
    pedido = df['Pedido'].values[lin]

    browser.get('https://portal.lenoxx.com.br/portal_lenoxx/sistema_forca_vendas/gerenciar_pedidos.php')

    numero_nf = browser.find_element_by_xpath('//*[@id="_easyui_textbox_input8"]')
    botao_procura = browser.find_element_by_xpath('//*[@id="tb"]/div[2]/a[1]')
    numero_nf.send_keys(nf)
    botao_procura.click()
    time.sleep(3)
    num_pedido = browser.find_element_by_xpath('/html/body/div[28]/div[2]/div[2]/div[2]/div[2]/table/tbody/tr[1]/td[1]/div').text



    tag = browser.find_element_by_xpath('/html/body/div[28]/div[2]/div[2]/div[2]/div[2]/table/tbody/tr[1]/td[9]/div/a[4]')
    tag = tag.get_attribute('href')
    browser.get(tag)
    time.sleep(3)

    numero_pedido = browser.find_elements_by_xpath('/html/body/div/div[2]/table/tbody/tr/td[1]/table/tbody/tr[3]')
    numero_pedido = numero_pedido[0].text
    numero_pedido = numero_pedido.replace('NÚMERO ', '')

    if num_pedido == numero_pedido:

        chkentrega = check_exists_by_xpath('/html/body/div/div[2]/table/tbody/tr/td[15]/table/tbody/tr[2]/td[2]/b')
        chkentrega2 = check_exists_by_xpath('/html/body/div/div[2]/table/tbody/tr/td[7]/table/tbody/tr[2]/td[2]/b')
        chkentrega3 = check_exists_by_xpath('/html/body/div/div[2]/table/tbody/tr[1]/td[15]/table/tbody/tr[2]/td[2]/b')
        chkentrega4 = check_exists_by_xpath('/html/body/div/div[2]/table/tbody/tr/td[7]/table/tbody/tr[2]/td[2]/b')

        if chkentrega == True:

            entrega = browser.find_element_by_xpath('/html/body/div/div[2]/table/tbody/tr/td[15]/table/tbody/tr[2]/td[2]/b').text

            if entrega == 'ENTREGA REGISTRADA':
                data_entrega = browser.find_element_by_xpath('/html/body/div/div[2]/table/tbody/tr/td[15]/table/tbody/tr[1]/td/i').text
                data_entrega = data_entrega.replace(' 00:00:00', '')
                status_entrega = 'Entregue'

            if entrega != 'ENTREGA REGISTRADA':
                entrega = ''
                data_entrega = ''
                status_entrega = 'Entrega não registrada no sistema'

            insert = """\
                    INSERT INTO PowerBiv2..Industrias_Data_Entrega_TEMP (FILIAL, Pedido, NF, SERIE_NF, Data, OBS)
                    VALUES(?, ?, ?, ?, FORMAT(CAST(? AS DATE), 'yyyyMMdd'), ?)
                    """
            status_entrega = 'Confirmação de entrega feita no Portal Lenoxx em ' + str(data_hora)

            cursor.execute(insert, '01', pedido, nf, serie_nf, data_entrega, status_entrega)



        if chkentrega2 == True:

            entrega = browser.find_element_by_xpath('/html/body/div/div[2]/table/tbody/tr/td[7]/table/tbody/tr[2]/td[2]/b').text

            if entrega == 'ENTREGA REGISTRADA':
                data_entrega = browser.find_element_by_xpath('/html/body/div/div[2]/table/tbody/tr/td[7]/table/tbody/tr[1]/td/i').text
                data_entreha = data_entrega.replace(' 00:00:00', '')
                status_entrega = 'Entregue'

            if entrega != 'ENTREGA REGISTRADA':
                entrega = ''
                data_entrega = ''
                status_entrega = 'Entrega não registrada no sistema'

            insert = """\
                    INSERT INTO PowerBiv2..Industrias_Data_Entrega_TEMP (FILIAL, Pedido, NF, SERIE_NF, Data, OBS)
                    VALUES(?, ?, ?, ?, FORMAT(CAST(? AS DATE), 'yyyyMMdd'), ?)
                    """
            status_entrega = 'Confirmação de entrega feita no Portal Lenoxx em ' + str(data_hora)

            cursor.execute(insert, '01', pedido, nf, serie_nf, data_entrega, status_entrega)

        else:
            entrega = ''
            data_entrega = ''
            status_entrega = 'Entrega não registrada no sistema'

            insert = """\
                            INSERT INTO PowerBiv2..Tabela_Consulta_Pedidos_Lenoxx(NF, SERIE_NF, PedidoIMPX, PedidoLenoxx, Entrega, DataEntrega, StatusEntrega, DataConsulta, HoraConsulta)
                            VALUES(?, ?, ?, ?, ?, ?, ?, ?, FORMAT(SYSDATETIME(), 'HH:mm'))
                      """
            cursor.execute(insert, nf, serie_nf, pedido, numero_pedido, entrega, data_entrega, status_entrega, hoje)


        if chkentrega3 == True:

            entrega = browser.find_element_by_xpath('/html/body/div/div[2]/table/tbody/tr[1]/td[15]/table/tbody/tr[2]/td[2]/b').text

            if entrega == 'ENTREGA REGISTRADA':
                data_entrega = browser.find_element_by_xpath('/html/body/div/div[2]/table/tbody/tr[1]/td[15]/table/tbody/tr[1]/td/i').text
                data_entreha = data_entrega.replace(' 00:00:00', '')
                status_entrega = 'Entregue'

                if entrega != 'ENTREGA REGISTRADA':
                    entrega = ''
                    data_entrega = ''
                    status_entrega = 'Entrega não registrada no sistema'

                insert = """\
                                INSERT INTO PowerBiv2..Industrias_Data_Entrega_TEMP (FILIAL, Pedido, NF, SERIE_NF, Data, OBS)
                                VALUES(?, ?, ?, ?, FORMAT(CAST(? AS DATE), 'yyyyMMdd'), ?)
                          """
                status_entrega = 'Verificação de entrega feita no Portal Lenoxx em ' + str(data_hora)

                cursor.execute(insert, '01', pedido, nf, serie_nf, data_entrega, status_entrega)

        else:
            entrega = ''
            data_entrega = ''
            status_entrega = 'Entrega não registrada no sistema'

            insert = """\
                            INSERT INTO PowerBiv2..Industrias_Data_Entrega_TEMP (FILIAL, Pedido, NF, SERIE_NF, Data, OBS)
                            VALUES(?, ?, ?, ?, FORMAT(CAST(? AS DATE), 'yyyyMMdd'), ?)
                      """
            status_entrega = 'Verificação de entrega feita no Portal Lenoxx em ' + str(data_hora)

            cursor.execute(insert, '01', pedido, nf, serie_nf, data_entrega, status_entrega)



        if chkentrega4 == True:

            entrega = browser.find_element_by_xpath('/html/body/div/div[2]/table/tbody/tr/td[7]/table/tbody/tr[2]/td[2]/b').text

            if entrega == 'ENTREGA REGISTRADA':
                data_entrega = browser.find_element_by_xpath('/html/body/div/div[2]/table/tbody/tr/td[7]/table/tbody/tr[1]/td/i').text
                data_entreha = data_entrega.replace(' 00:00:00', '')
                status_entrega = 'Entregue'

                if entrega != 'ENTREGA REGISTRADA':
                    entrega = ''
                    data_entrega = ''
                    status_entrega = 'Entrega não registrada no sistema'

                insert = """\
                                INSERT INTO PowerBiv2..Industrias_Data_Entrega_TEMP (FILIAL, Pedido, NF, SERIE_NF, Data, OBS)
                                VALUES(?, ?, ?, ?, FORMAT(CAST(? AS DATE), 'yyyyMMdd'), ?)
                          """
                status_entrega = 'Verificação de entrega feita no Portal Lenoxx em ' + str(data_hora)

                cursor.execute(insert, '01', pedido, nf, serie_nf, data_entrega, status_entrega)


        else:
            entrega = ''
            data_entrega = ''
            status_entrega = 'Entrega não registrada no sistema'

            insert = """\
                            INSERT INTO PowerBiv2..Industrias_Data_Entrega_TEMP (FILIAL, Pedido, NF, SERIE_NF, Data, OBS)
                            VALUES(?, ?, ?, ?, FORMAT(CAST(? AS DATE), 'yyyyMMdd'), ?)
                      """
            cursor.execute(insert, nf, serie_nf, pedido, numero_pedido, entrega, data_entrega, status_entrega, hoje)

    lin = lin + 1

insert = """\
                INSERT INTO PJKTB2..ZX4120 (ZX4_FILIAL, ZX4_PEDIDO, ZX4_NF, ZX4_SERIE, ZX4_STATUS, ZX4_DATA, ZX4_OBS, R_E_C_N_O_)
					SELECT  
							FILIAL,
							RTRIM(N.Pedido), 
							N.NF, 
							N.SERIE_NF,                 
							'60' Status,
							Data,
							OBS,
							(SELECT COALESCE(MAX(PJKTB2..ZX4120.R_E_C_N_O_),0) FROM PJKTB2..ZX4120) + 
							(ROW_NUMBER() OVER(ORDER BY N.NF)) R_E_C_N_O_  
					FROM PowerBiv2..Industrias_Data_Entrega_TEMP N
					LEFT JOIN PJKTB2..ZX4120 X ON X.ZX4_NF = N.NF AND ZX4_PEDIDO = N.PEDIDO AND ZX4_STATUS = '60'
					WHERE ZX4_STATUS IS NULL
					GROUP BY  
							N.FILIAL,
							N.NF, 
							N.SERIE_NF,  
							N.Pedido,                
							Data,
							OBS
          """
cursor.execute(insert)
cursor.execute('EXEC DataWareHouse.dbo.IndustriaPedidoNFe')

browser.quit()