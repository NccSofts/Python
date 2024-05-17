import os
import pandas as pd
import time
import pyodbc
import sqlalchemy
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
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
cursor = db.sql_conn('PowerBiv2')
engine = db.sql_engine('PowerBiv2')

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
            WHERE Industria = 'DL' AND Status1 IN ('Pendente Entrega' , 'Pendente Entrega em atraso') AND Nome_Tranportadora = 'SOLISTICA'
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
                WHERE Industria = 'DL' AND Status1 IN ('Pendente Entrega' , 'Pendente Entrega em atraso') AND Nome_Tranportadora = 'SOLISTICA'
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

    popup = check_exists_by_xpath('/html/body/div/div/div[3]/button[1]')
    if popup == True:
        popup = browser.find_element_by_xpath('/html/body/div/div/div[3]/button[1]')
        popup.get_attribute('href')
        popup.click()


    industria = df['Industria'].values[lin]
    nf = df['NF'].values[lin]
    serie_nf = df['SERIE_NF'].values[lin]
    pedido = df['PEDIDO'].values[lin]
    cnpj_destino = df['CNPJ'].values[lin]
    transportadora = df['Nome_Tranportadora'].values[lin]

    # --> VERIFICA SE O CAMPO TIPO CLIENTE ESTA NA TELA



    chkTipoDocumento = check_exists_by_xpath('/html/body/app-root/app-login/div/div/div[1]/div/div/form/div[1]/div[1]/div/ng-select/div/div/div[2]/input')
    if chkTipoDocumento == True:
        tipocliente = browser.find_element_by_xpath('/html/body/app-root/app-login/div/div/div[1]/div/div/form/div[1]/div[1]/div/ng-select/div/div/div[2]/input')
        tipocliente.send_keys('N')
        tipocliente.send_keys(Keys.ENTER)
        time.sleep(1)
        nf_site = browser.find_element_by_xpath(
            '/html/body/app-root/app-login/div/div/div[1]/div/div/form/div[1]/div[2]/div/input')
        nf_site.send_keys(nf)

    chkCNPJ_CPF = check_exists_by_id("cnpjCpf")
    if chkCNPJ_CPF == True:
        cnpj = browser.find_element_by_id("cnpjCpf")
        cnpj.send_keys(cnpj_destino)


    chkBotaoConsulta = check_exists_by_xpath("/html/body/app-root/app-login/div/div/div[1]/div/div/form/div[2]/div[2]/div/button")
    # chkBotaoConsulta = check_exists_by_id("buscar")
    if chkBotaoConsulta == True:
        botaoConsulta = browser.find_element_by_xpath("/html/body/app-root/app-login/div/div/div[1]/div/div/form/div[2]/div[2]/div/button")
        botaoConsulta.get_attribute('href')
        botaoConsulta.click()

    # time.sleep(1)
    now = datetime.now()
    date_time = now.strftime("%d/%m/%Y - %H:%M")

    try:
        btn_nf_consulta = WebDriverWait(browser, 30).until(EC.element_to_be_clickable((By.XPATH,
                    "/html/body/div/div/div[2]/div/app-consulta-publica/div/div[3]/div/ag-grid-angular/div/div[1]/div/div[3]/div[2]/div/div/div/div[1]")))
        btn_nf_consulta.click()
    except:
        status = "Não Localizada"
        obs = "NF não encontrada no site da transportadora " + transportadora + " em " + str(date_time)


    try:
        previsao_entrega = browser.find_element_by_xpath(
            '/html/body/div/div/div[2]/div/app-consulta-publica/div/div[4]/div[5]/div[2]').text
    except:
        previsao_entrega = None

    try:
        data_entrega = browser.find_element_by_xpath(
            '/html/body/div/div/div[2]/div/app-consulta-publica/div/div[4]/div[6]/div[2]').text
        if data_entrega == '':
            data_entrega = None
    except:
        data_entrega = None
        obs = "Entrega verificada no site da transportadora " + transportadora + " em " + str(date_time)

    try:
        cte_transp = browser.find_element_by_xpath(
            'html/body/div/div/div[2]/div/app-consulta-publica/div/div[4]/div[1]/div[2]').text
    except:
        cte_transp = None

    if data_entrega != None:
        status = "Entregue"
        obs = "Entrega confirmada no site da transportadora " + transportadora + " em " + str(date_time)

    if data_entrega == None:
        data_entrega = previsao_entrega
        status = "Entrega ainda não realizada"
        obs = "Entrega verificada no site da transportadora " + transportadora + " em " + str(date_time)



    insert = """\
                    INSERT INTO PowerBiv2..Tabela_Consulta_Pedidos_DL(NF, SERIE_NF, Pedido_IMPX, StatusEntrega, Data_Entrega, DataConsulta, HoraConsulta, SituacaoEntrega)
                    VALUES(?, ?, ?, ?, ?, ?, ?, ?)
              """
    cursor.execute(insert, nf, serie_nf, pedido, obs, data_entrega, str(now.strftime("%d/%m/%Y")),
                   str(now.strftime("%H:%M")), status)

    lin = lin + 1

browser.close()


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

