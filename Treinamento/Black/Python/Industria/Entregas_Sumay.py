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
# from Database.database import sql_conn, sql_engine, hoje



dataatual = datetime.datetime.now()
diasemana = datetime.datetime.today().weekday()
hora_atual = str(datetime.datetime.today().time())

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


# Definir conexao 0 = PowerBiV2 / 1 = Datawarehouse
cursor = db.sql_conn('PowerBiv2')
engine = db.sql_engine('PowerBiv2')


lista = """\
                SELECT  
                        Industria,
                        NF, 
                        SERIE_NF,  
                        Pedido,                
                        CNPJ  
                
                FROM PowerBIv2..Acompanhamento_Industrias_Base_v2
                WHERE Data_Pedido >= '2021-01-01' AND Industria = 'Sumay' AND Status1 IN ('Pendente Entrega' , 'Pendente Entrega em atraso')
                --WHERE Pedido = '1337'
                GROUP BY  
                        Industria,
                        NF,
                        CNPJ,
                        Pedido,
                        SERIE_NF
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



for items in df.itertuples():
    # ABRE BROWSER
    browser.get('https://portalsmo.mosistemas.com/track/rastreamento/#/')

    # DADOS PARA CONSULTA DE ENTREGA
    #----> AGUARDA A PAGINA CARREGAR
    time.sleep(3)

    industria = df['Industria'].values[lin]
    nf = df['NF'].values[lin]
    serie_nf = df['SERIE_NF'].values[lin]
    cnpj_destino = df['CNPJ'].values[lin]
    pedido = df['Pedido'].values[lin]
    filial = '01'
    status = '60'
    obs = 'Entrega confirmada pelo site de rastreio da Sumay em ' + str(hoje) + ' / Lançamento automatico'

    # --> VERIFICA SE O CAMPO TIPO CLIENTE ESTA NA TELA
    chkTipoCliente =check_exists_by_id("input-tipoCliente")
    if chkTipoCliente == True:
        tipocliente = browser.find_element_by_id("input-tipoCliente")
        tipocliente.send_keys('D')

    chkCNPJ_CPF = check_exists_by_id("input-cnpjCpf")
    if chkCNPJ_CPF == True:
        cnpj = browser.find_element_by_id("input-cnpjCpf")
        cnpj.send_keys(cnpj_destino)

    chkTipoDocumento = check_exists_by_id("input-tipoDocumento")
    if chkTipoDocumento == True:
        tipoDocumento = browser.find_element_by_id("input-tipoDocumento")
        tipoDocumento.send_keys('N')

    chkNumeroDocumento = check_exists_by_id("input-numeroDoDocumento")
    if chkNumeroDocumento == True:
        numeroDocumento = browser.find_element_by_id("input-numeroDoDocumento")
        numeroDocumento.send_keys(nf)

    chkBotaoConsulta = check_exists_by_xpath('//*[@id="form"]/div/div[1]/div[5]/div/div')
    if chkBotaoConsulta == True:
        botaoConsulta = browser.find_element_by_xpath('//*[@id="form"]/div/div[1]/div[5]/div/div')
        botaoConsulta.click()

    time.sleep(10)

    chkemissao = check_exists_by_xpath('//*[@id="form"]/div/div/div[3]/div/div[2]/div[4]')
    if chkemissao == True:
        emissaoNF = browser.find_element_by_xpath('//*[@id="form"]/div/div/div[3]/div/div[2]/div[4]')
        dataEmissaoNF = emissaoNF.text

    chkEntrega = check_exists_by_xpath('//*[@id="form"]/div/div/div[3]/div/div[2]/div[7]')
    if chkEntrega == True:
        Entrega = browser.find_element_by_xpath('//*[@id="form"]/div/div/div[3]/div/div[2]/div[7]')
        posicaoEntrega = str(Entrega.text)
        posicaoEntrega = posicaoEntrega.replace('\n\n', ' ')
        posicaoEntrega = posicaoEntrega.replace('\n', ' ')

        teste_entregue = posicaoEntrega.find("Entregue")
        teste_previsão = posicaoEntrega.find("Previsão de Entrega")


        if teste_entregue == 0:
            Status_Entrega = 'Entregue'
            dataEntrega = str(substring(posicaoEntrega, 9, 14) + '/' + dataatual.strftime("%Y"))

        if teste_previsão == 0:
            Status_Entrega = 'Previsão de Entrega'
            dataEntrega = str(substring(posicaoEntrega, 20, 31))


    chkOcorrencias = check_exists_by_xpath('//*[@id="form"]/div/div/div[3]/div/div[4]/div')
    if chkOcorrencias == True:
        ocorrencias = browser.find_element_by_xpath('//*[@id="form"]/div/div/div[3]/div/div[4]/div')
        ocorrenciasEntrega = str(ocorrencias.text)
        ocorrenciasEntrega = ocorrenciasEntrega.replace('\n', ' ')
    else:
        ocorrenciasEntrega = ''

    insert = """\
                    INSERT INTO PowerBIv2..Status_Entregas_Industrias(Industria, NF, CNPJ, StatusEntrega, Data_Entrega, Ocorrencias)
                    VALUES(?, ?, ?, ?, ?, ?)
              """
    cursor.execute(insert, industria, nf, cnpj_destino, Status_Entrega, dataEntrega, ocorrenciasEntrega)

    # if lin + 1 <= len(df):
    lin = lin + 1

previsao = """\
                UPDATE PJKTB2..ZX1110          
                    SET ZX1_PENTR = FORMAT(CAST(S.Data_Entrega AS DATE),  'yyyyMMdd')              
                FROM PJKTB2..ZX1110 Z
                INNER JOIN PowerBIv2..Status_Entregas_Industrias S ON S.NF = Z.ZX1_NF AND ZX1_CNPJ = S.CNPJ
                            AND S.Industria = 'Sumay' AND S.Statusentrega = 'Previsão de Entrega'
                WHERE ZX1_PENTR <> FORMAT(CAST(S.Data_Entrega AS DATE),  'yyyyMMdd')
        """

cursor.execute(previsao)


entrega = """\
                DELETE FROM PowerBIv2..Status_Entregas_Industrias
                WHERE LEN(Data_Entrega) < 10
                
                INSERT INTO PJKTB2..ZX4110 (ZX4_FILIAL, ZX4_PEDIDO, ZX4_NF, ZX4_SERIE, ZX4_STATUS, ZX4_DATA, ZX4_OBS, R_E_C_N_O_)
        
                SELECT  
                        '01' FILIAL,
                        N.Pedido, 
                        N.NF, 
                        N.SERIE_NF,                 
                        '60' Status,
                        FORMAT(CAST(S.Data_Entrega AS DATE), 'yyyyMMdd') Data,
                        CONCAT('Entrega confirmada pelo site de rastreio da Sumay em ', FORMAT(GETDATE(),'dd-MM-yyyy') ,' / Lançamento automatico') OBS,
                        (SELECT COALESCE(MAX(PJKTB2..ZX4110.R_E_C_N_O_),0) FROM PJKTB2..ZX4110) + 
                        (ROW_NUMBER() OVER(ORDER BY N.NF)) R_E_C_N_O_  
                        --(ZX4_FILIAL, ZX4_PEDIDO, ZX4_NF, ZX4_SERIE, ZX4_STATUS, ZX4_DATA, ZX4_OBS)
                FROM PowerBIv2..Acompanhamento_Industrias_Base_v2 N
                INNER JOIN PowerBIv2..Status_Entregas_Industrias S ON S.Industria = N.Industria AND S.NF = N.NF AND S.StatusEntrega = 'Entregue'
                LEFT JOIN PJKTB2..ZX4110 X ON X.ZX4_NF = N.NF AND ZX4_PEDIDO = N.PEDIDO AND ZX4_STATUS = '60'
                WHERE ZX4_STATUS IS NULL
                GROUP BY  
                        N.Industria,
                        N.NF, 
                        N.SERIE_NF,  
                        N.Pedido,                
                        N.CNPJ,
                        S.Data_Entrega
        """
cursor.execute(entrega)

cursor.execute('EXEC DataWareHouse.dbo.IndustriaPedidoNFe')


# browser.close()
browser.quit()


