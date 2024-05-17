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
from selenium.webdriver.common.keys import Keys
from datetime import timedelta, date
from Database.database import sql_conn, sql_engine
import time

dataatual = datetime.datetime.now()
diasemana = datetime.datetime.today().weekday()
hora_atual = str(datetime.datetime.today().time())


def left(s, amount):
    return s[:amount]

def right(s, amount):
    return s[-amount:]

def mid(s, offset, amount):
    return s[offset:offset+amount]



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



# Definir conexao 0 = PowerBiV2 / 1 = Datawarehouse
cursor = sql_conn('PowerBiv2')
engine = sql_engine('PowerBiv2')

# Data de hoje
hoje = date.today().strftime('%d/%m/%Y')
datai = date.today()
data_query = date.today().strftime('%Y%m%d')


lista = """\
                SELECT [CNPJ Transportadora] CNPJ_Transp
                FROM PowerBIv2..Acompanhamento_Industrias_Base_v2 N
                LEFT JOIN PowerBiv2..Cadastro_Transportadoras_Industrias T ON RTRIM(T.CNPJ) = RTRIM([CNPJ Transportadora])
                WHERE [CNPJ Transportadora] <> '' AND T.Razao_Social IS NULL
                GROUP BY [CNPJ Transportadora]

          """
df = pd.read_sql(lista, engine)
lin = 0

# Conexao HOST Selenium

chrome_options = webdriver.ChromeOptions();
chrome_options.add_experimental_option("excludeSwitches", ['enable-logging']);
# browser = webdriver.Chrome(options=chrome_options);

browser = webdriver.Chrome(executable_path=r'C:/Python/Selenium/chromedriver.exe',
                           options=chrome_options)  # Optional argument, if not specified will search path
browser.maximize_window()
action = ActionChains(browser)

lin = 0
for items in df.itertuples():
    cnpj = df['CNPJ_Transp'].values[lin]
    lin = lin + 1

    # ABRE BROWSER
    browser.get('https://cnpj.biz/' + cnpj)

    chkCNPJ = check_exists_by_xpath('/ html / body / div[1] / div / div / div[1] / p[1] / span[2] / b')
    if chkCNPJ == True:

        try:
            cnpj_transp = browser.find_element_by_xpath('/ html / body / div[1] / div / div / div[1] / p[1] / span[2] / b').text
        except:
            cnpj_transp = ''

        try:
            cnpj_transp_formatado = browser.find_element_by_xpath('/ html / body / div[1] / div / div / div[1] / p[1] / span[1] / b').text
        except:
            cnpj_transp_formatado = ''


        try:
            nome_fantasia_transp = mid(browser.find_element_by_xpath("//*[contains(text(), 'Nome Fantasia:')]").text , 15,
                                      len(browser.find_element_by_xpath("//*[contains(text(), 'Nome Fantasia:')]").text))
        except:
            nome_fantasia_transp = ''


        try:
            razao_social_transp = mid(browser.find_element_by_xpath("//*[contains(text(), 'Razão Social:')]").text , 14,
                                      len(browser.find_element_by_xpath("//*[contains(text(), 'Razão Social:')]").text))
        except:
            razao_social_transp = ''

        try:
            endereco = mid(browser.find_element_by_xpath("//*[contains(text(), 'Logradouro:')]").text , 11,
                                      len(browser.find_element_by_xpath("//*[contains(text(), 'Logradouro:')]").text))
        except:
            endereco = ''


        try:
            complemento = mid(browser.find_element_by_xpath("//*[contains(text(), 'Complemento:')]").text , 13,
                                      len(browser.find_element_by_xpath("//*[contains(text(), 'Complemento:')]").text))
        except:
            complemento = ''

        try:
            bairro = mid(browser.find_element_by_xpath("//*[contains(text(), 'Bairro:')]").text , 7,
                                      len(browser.find_element_by_xpath("//*[contains(text(), 'Bairro:')]").text))
        except:
            bairro = ''


        try:
            cep = mid(browser.find_element_by_xpath("//*[contains(text(), 'CEP:')]").text , 4,
                                      len(browser.find_element_by_xpath("//*[contains(text(), 'CEP:')]").text))
        except:
            cep = ''


        try:
            municipio = mid(browser.find_element_by_xpath("//*[contains(text(), 'Município:')]").text , 11,
                                      len(browser.find_element_by_xpath("//*[contains(text(), 'Município:')]").text))
        except:
            municipio = ''

        try:
            estado = mid(browser.find_element_by_xpath("//*[contains(text(), 'Estado:')]").text , 8,
                                      len(browser.find_element_by_xpath("//*[contains(text(), 'Estado:')]").text))
        except:
            estado = ''

        insert = """\
                        INSERT INTO PowerBiV2..Cadastro_Transportadoras_Industrias
                        VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?)
                  """
        cursor.execute(insert, str(cnpj), str(razao_social_transp.upper()), str(nome_fantasia_transp.upper()), str(endereco.upper()), str(complemento.upper()), str(bairro.upper()), str(cep.upper())
                       , str(municipio.upper()), str(estado.upper()))


    # DADOS PARA CONSULTA DE ENTREGA
    # ----> AGUARDA A PAGINA CARREGAR
    # time.sleep(3)


browser.quit()