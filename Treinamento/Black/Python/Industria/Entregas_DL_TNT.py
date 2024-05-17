import os
import pandas as pd
import time
import pyodbc
import sqlalchemy
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support import wait as WebDriverWait
from selenium.webdriver.common.keys import Keys
from datetime import timedelta, date
from datetime import datetime
import time
import sys
sys.path.append('C:/Python/Database/')
import database as db



dataatual = datetime.now()
diasemana = datetime.today().weekday()
hora_atual = str(datetime.today().time())

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


def check_exists_by_class(name):
    try:
        browser.find_element_by_class_name(name)
    except NoSuchElementException:
        return False
    return True


def check_exists_by_id(name):
    try:
        browser.find_element_by_id(name)
    except NoSuchElementException:
        return False
    return True


def check_exists_by_link_text(name):
    try:
        browser.find_element_by_link_text(name)
    except NoSuchElementException:
        return False
    return True


def frame_switch(name):
  browser.switch_to.frame(browser.find_element_by_name(name))

def substring(s, start, end):
    return s[start:end]

# Data de hoje
hoje = date.today().strftime('%d/%m/%Y')


# Definir conexao 0 = PowerBiV2 / 1 = Datawarehouse
opcao = 0

configuracoes = pd.read_csv("C:/Python/database.cfg", header=0, delimiter=";")

server = configuracoes['IP'].values[opcao]
database = configuracoes['Banco'].values[opcao]
username = configuracoes['Usuario'].values[opcao]
password = configuracoes['Senha'].values[opcao]

sql_string = 'mssql+pyodbc://' + username + ':' + password + '@' + server + '/' + database + '?driver=ODBC+Driver+13+for+SQL+Server'

cnxn = pyodbc.connect(
    'DRIVER={ODBC Driver 13 for SQL Server};SERVER=' + server + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password,
    autocommit=True)
cursor = cnxn.cursor()

status_check_cursor = pyodbc.connect(
    'DRIVER={ODBC Driver 13 for SQL Server};SERVER=' + server + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password,
    autocommit=True)
cursor2 = status_check_cursor.cursor()

# Conexao com o servidor SQL
engine = sqlalchemy.create_engine(sql_string)
engine.connect()

lista = """\
            SELECT  
                    Industria,
                    NF,
                    SERIE_NF, 
                    PEDIDO,                  
                    CNPJ,
                    Nome_Tranportadora,
                    WEBSITE
            FROM PowerBIv2..Acompanhamento_Industrias_Base_v2
            WHERE Industria = 'DL' AND Status1 IN ('Pendente Entrega' , 'Pendente Entrega em atraso') AND Nome_Tranportadora = 'TNT'
            GROUP BY  
                    Industria,
                    NF,
                    SERIE_NF,
                    PEDIDO,
                    CNPJ,
                    Nome_Tranportadora,
                    WEBSITE
          """
df = pd.read_sql(lista, engine)
lin = 0

cursor.execute('TRUNCATE TABLE [PowerBIv2].[dbo].[Status_Entregas_Industrias]')

#Conexao HOST Selenium
chrome_options = webdriver.ChromeOptions();
chrome_options.add_experimental_option("excludeSwitches", ['enable-logging']);

browser = webdriver.Chrome(executable_path=r'C:/Python/Selenium/chromedriver.exe', options=chrome_options)  # Optional argument, if not specified will search path
browser.maximize_window()
action = ActionChains(browser)


delete = """\
                DELETE w FROM PowerBiv2..Tabela_Consulta_Pedidos_DL w
                INNER JOIN
                (
                SELECT  
                        Industria,
                        NF,
                        SERIE_NF, 
                        PEDIDO,                  
                        CNPJ,
                        Nome_Tranportadora,
                        WEBSITE
                FROM PowerBIv2..Acompanhamento_Industrias_Base_v2
                WHERE Industria = 'DL' AND Status1 IN ('Pendente Entrega' , 'Pendente Entrega em atraso') AND Nome_Tranportadora = 'TNT'
                GROUP BY  
                        Industria,
                        NF,
                        SERIE_NF,
                        PEDIDO,
                        CNPJ,
                        Nome_Tranportadora,
                        WEBSITE
                ) AS B ON B.NF = W.NF
          """
cursor.execute(delete)


for items in df.itertuples():

    website = df['WEBSITE'].values[lin]

    # ABRE BROWSER
    browser.get(website)

    # DADOS PARA CONSULTA DE ENTREGA
    #----> AGUARDA A PAGINA CARREGAR
    time.sleep(3)

    industria = df['Industria'].values[lin]
    nf = df['NF'].values[lin]
    serie_nf = df['SERIE_NF'].values[lin]
    pedido = df['PEDIDO'].values[lin]
    cnpj_destino = df['CNPJ'].values[lin]
    transportadora = df['Nome_Tranportadora'].values[lin]

    # --> VERIFICA SE O CAMPO TIPO CLIENTE ESTA NA TELA
    chkTipoCliente =check_exists_by_id("remDest")
    if chkTipoCliente == True:
        tipocliente = browser.find_element_by_id("remDest")
        tipocliente.send_keys('D')

    chkCNPJ_CPF = check_exists_by_id("nrIdentificacao")
    if chkCNPJ_CPF == True:
        cnpj = browser.find_element_by_id("nrIdentificacao")
        cnpj.send_keys(cnpj_destino)

    chkTipoDocumento = check_exists_by_id("tpDocumento")
    if chkTipoDocumento == True:
        tipoDocumento = browser.find_element_by_id("tpDocumento")
        tipoDocumento.send_keys('N')

    chkNumeroDocumento = check_exists_by_id("nrDocumento")
    if chkNumeroDocumento == True:
        numeroDocumento = browser.find_element_by_id("nrDocumento")
        numeroDocumento.send_keys(nf)



    chkBotaoConsulta = check_exists_by_xpath("/html/body/div/section/div/div/div/form[1]/div/div/div[6]/a[1]")
    # chkBotaoConsulta = check_exists_by_id("buscar")
    if chkBotaoConsulta == True:
        botaoConsulta = browser.find_element_by_xpath("/html/body/div/section/div/div/div/form[1]/div/div/div[6]/a[1]")
        botaoConsulta.get_attribute('href')
        botaoConsulta.click()

    time.sleep(3)

    # while chkNfConsulta == False:
    chkNfConsulta = check_exists_by_xpath("/html/body/div/section/div/div/div/form[2]/div/div/div/table/tbody/tr/td[1]/a")



    if chkNfConsulta == True:
        nf_consultar = browser.find_element_by_xpath("/html/body/div/section/div/div/div/form[2]/div/div/div/table/tbody/tr/td[1]/a").get_attribute("href")
        browser.get(nf_consultar)

    time.sleep(3)

    chk_entrega = check_exists_by_xpath("/html/body/div[1]/header[2]/div[2]/div/div/div/div[3]/div[5]")

    if chk_entrega == False:
        status = "Não Localizada"
        now = datetime.now()
        date_time = now.strftime("%d/%m/%Y - %H:%M")
        obs = "NF não encontrada no site da transportadora " + transportadora + " em " + str(date_time)
        data_entrega = None

        insert = """\
                        INSERT INTO PowerBiv2..Tabela_Consulta_Pedidos_DL(NF, SERIE_NF, Pedido_IMPX, StatusEntrega, Data_Entrega, DataConsulta, HoraConsulta, SituacaoEntrega)
                        VALUES(?, ?, ?, ?, ?, ?, ?, ?)
                  """
        cursor.execute(insert, nf, serie_nf, pedido, obs, data_entrega, str(now.strftime("%d/%m/%Y")), str(now.strftime("%H:%M")), status)


    if chk_entrega == True:
        entrega = db.left(browser.find_element_by_xpath("/html/body/div[1]/header[2]/div[2]/div/div/div/div[3]/div[5]").text, 19)

        if entrega == "PREVISÃO DE ENTREGA":
            status = "Entrega ainda não realizada"
            now = datetime.now()
            date_time = now.strftime("%d/%m/%Y - %H:%M")
            obs = "Entrega verificada no site da transportadora " + transportadora + " em " + str(date_time)
            previsao_entrega = browser.find_element_by_xpath("/html/body/div[1]/header[2]/div[2]/div/div/div/div[3]/div[5]/div[2]/div[1]/div").text
            data_entrega = previsao_entrega

            insert = """\
                            INSERT INTO PowerBiv2..Tabela_Consulta_Pedidos_DL(NF, SERIE_NF, Pedido_IMPX, StatusEntrega, Data_Entrega, DataConsulta, HoraConsulta, SituacaoEntrega)
                            VALUES(?, ?, ?, ?, ?, ?, ?, ?)
                      """
            cursor.execute(insert, nf, serie_nf, pedido, obs, data_entrega, str(now.strftime("%d/%m/%Y")),
                           str(now.strftime("%H:%M")), status)


    if chk_entrega == True:
        entrega = db.left(browser.find_element_by_xpath("/html/body/div[1]/header[2]/div[2]/div/div/div/div[3]/div[5]").text, 10)

        if entrega == "FINALIZADO":
            status = "Entregue"
            now = datetime.now()
            date_time = now.strftime("%d/%m/%Y - %H:%M")
            obs = "Entrega confirmada no site da transportadora " + transportadora + " em " + str(date_time)
            data_entrega = db.left(browser.find_element_by_xpath("/html/body/div[1]/header[2]/div[2]/div/div/div/div[3]/div[5]/div[2]/div[1]/div").text, 10)

        insert = """\
                        INSERT INTO PowerBiv2..Tabela_Consulta_Pedidos_DL(NF, SERIE_NF, Pedido_IMPX, StatusEntrega, Data_Entrega, DataConsulta, HoraConsulta, SituacaoEntrega)
                        VALUES(?, ?, ?, ?, ?, ?, ?, ?)
                  """
        cursor.execute(insert, nf, serie_nf, pedido, obs, data_entrega, str(now.strftime("%d/%m/%Y")), str(now.strftime("%H:%M")), status)

    lin = lin + 1



previsao = """\
                UPDATE PJKTB2..ZX1080         
                    SET ZX1_PENTR = FORMAT(CAST(P.Data_Entrega AS DATE),  'yyyyMMdd')              
                FROM PJKTB2..ZX1080 Z
                INNER JOIN PowerBiv2..Tabela_Consulta_Pedidos_DL P ON P.NF = ZX1_NF 
                        AND  P.SituacaoEntrega = 'Entrega ainda não realizada'
                        AND P.Data_Entrega IS NOT NULL
                WHERE ZX1_PENTR <> FORMAT(CAST(P.Data_Entrega AS DATE),  'yyyyMMdd')
            """

cursor.execute(previsao)

entrega = """\
                INSERT INTO PJKTB2..ZX4080 (ZX4_FILIAL, ZX4_PEDIDO, ZX4_NF, ZX4_SERIE, ZX4_STATUS, ZX4_DATA, ZX4_OBS, R_E_C_N_O_)

                SELECT  
                        '01' FILIAL,
                        N.Pedido, 
                        N.NF, 
                        N.SERIE_NF,                 
                        '60' Status,
                        FORMAT(CAST(P.Data_Entrega AS DATE), 'yyyyMMdd') Data,
                        CONCAT(P.StatusEntrega,' / Lançamento automatico') OBS,
                        (SELECT COALESCE(MAX(PJKTB2..ZX4080.R_E_C_N_O_),0) FROM PJKTB2..ZX4080) + 
                        (ROW_NUMBER() OVER(ORDER BY N.NF)) R_E_C_N_O_  
                FROM PowerBIv2..Acompanhamento_Industrias_Base_v2 N
                INNER JOIN PowerBiv2..Tabela_Consulta_Pedidos_DL P ON P.NF = N.NF AND P.SituacaoEntrega = 'Entregue'
                LEFT JOIN PJKTB2..ZX4080 X ON X.ZX4_NF = N.NF AND ZX4_PEDIDO = N.PEDIDO AND ZX4_STATUS = '60'
                WHERE ZX4_STATUS IS NULL
                GROUP BY  
                        N.Industria,
                        N.NF, 
                        N.SERIE_NF,  
                        N.Pedido,                
                        N.CNPJ,
                        P.Data_Entrega,
                        P.StatusEntrega
        """
cursor.execute(entrega)

cursor.execute('EXEC DataWareHouse.dbo.IndustriaPedidoNFe')

browser.quit()


