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
cursor = db.sql_conn('PowerBiv2')
engine = db.sql_engine('PowerBiv2')

lista = """\
             INSERT INTO PowerBiv2..Tabela_Entregas_MP (Numero, Transportadora)
            
            SELECT 
                R.Numero,
                R.Transportadora
            
            FROM PowerBIv2..Tabela_Rentabilidade_BI R 
            LEFT JOIN PowerBiv2..Tabela_Entregas_MP E ON E.Numero = R.Numero
            WHERE E.Numero IS NULL AND Ano = '2021' AND Mes >= '11' AND TipoNF = 'Venda' 
                AND R.Transportadora = 'TNT' AND R.DataEntrega IS NULL
            GROUP BY 	
                        R.Numero, R.Transportadora
            ORDER BY R.Numero
          """
cursor.execute(lista)


lista = """\
                SELECT 
                    M.Numero,
                    REPLACE(M.Numero, '.3', '')*1 NF,
                    C.CNPJ,
                    C.Transportadora
                FROM PowerBiv2..Tabela_Entregas_MP M
                INNER JOIN
                (
                    SELECT 
                        Numero,
                        CNPJ,
                        Transportadora
                    FROM PowerBIv2..Tabela_Rentabilidade_BI R
                    GROUP BY Numero, CNPJ, Transportadora
                ) C ON C.Numero = M.Numero AND Entrega IS NULL
                ORDER BY Numero
          """
df = pd.read_sql(lista, engine)
lin = 0

#Conexao HOST Selenium
browser = webdriver.Chrome(executable_path=r'C:/Python/Selenium/chromedriver.exe')  # Optional argument, if not specified will search path
browser.maximize_window()
action = ActionChains(browser)


for items in df.itertuples():

    data_entrega = None
    previsao_entrega = None
    status = None
    emitido = None

    website = 'https://radar.tntbrasil.com.br/radar/public/localizacaoSimplificada.do'

    # ABRE BROWSER
    browser.get(website)

    # DADOS PARA CONSULTA DE ENTREGA
    #----> AGUARDA A PAGINA CARREGAR
    time.sleep(3)

    nf = str(df['NF'].values[lin])
    numero_nf = str(df['Numero'].values[lin])
    cnpj_destino = df['CNPJ'].values[lin]
    transportadora = df['Transportadora'].values[lin]

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
    chk_emitido = check_exists_by_xpath('/html/body/div/header[2]/div[2]/div/div/div/div[3]/div[1]/div[1]')




    if chk_emitido == True:
        emitido = db.left(browser.find_element_by_xpath('/html/body/div/header[2]/div[2]/div/div/div/div[3]/div[1]/div[2]').text, 10)
    else:
        emitido = None


    if chk_entrega == True:
        entrega = db.left(browser.find_element_by_xpath("/html/body/div[1]/header[2]/div[2]/div/div/div/div[3]/div[5]").text, 19)

        if entrega == "PREVISÃO DE ENTREGA":
            status = "Entrega ainda não realizada"
            now = datetime.now()
            date_time = now.strftime("%d/%m/%Y - %H:%M")
            obs = "Entrega verificada no site da transportadora " + transportadora + " em " + str(date_time)
            previsao_entrega = browser.find_element_by_xpath("/html/body/div[1]/header[2]/div[2]/div/div/div/div[3]/div[5]/div[2]/div[1]/div").text
            data_entrega = None

        entrega = db.left(browser.find_element_by_xpath("/html/body/div[1]/header[2]/div[2]/div/div/div/div[3]/div[5]").text, 10)

        if entrega == "FINALIZADO":
            status = "Entregue"
            now = datetime.now()
            date_time = now.strftime("%d/%m/%Y - %H:%M")
            obs = "Entrega confirmada no site da transportadora " + transportadora + " em " + str(date_time)
            data_entrega = db.left(browser.find_element_by_xpath("/html/body/div[1]/header[2]/div[2]/div/div/div/div[3]/div[5]/div[2]/div[1]/div").text, 10)


    insert = """\
                        UPDATE PowerBiv2..Tabela_Entregas_MP
                            SET Expedicao = iif(? IS NULL, Expedicao, ?),  
                                PrevisaoEntrega = iif(? IS NULL, PrevisaoEntrega, ?), 
                                Entrega = iif(CAST(? AS DATE) IS NULL, Entrega,CAST(? AS DATE)) , StatusAtual = ?                      
                        WHERE Numero = ?
              """
    cursor.execute(insert, emitido, emitido, previsao_entrega, previsao_entrega, data_entrega, data_entrega, status, numero_nf)

    lin = lin + 1


limpar = """\
                UPDATE PowerBiv2..Tabela_Entregas_MP
                    SET Expedicao = NULL
                WHERE Expedicao = '1900-01-01'
                
                UPDATE PowerBiv2..Tabela_Entregas_MP
                    SET PrevisaoEntrega = NULL
                WHERE PrevisaoEntrega = '1900-01-01'
                
                UPDATE PowerBiv2..Tabela_Entregas_MP
                    SET Entrega = NULL
                WHERE Entrega = '1900-01-01'
          """
cursor.execute(limpar)

cursor.execute('EXEC DataWareHouse.dbo.IndustriaPedidoNFe')

browser.quit()


