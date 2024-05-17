import os
import pandas as pd
from openpyxl import load_workbook
import math
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.keys import  Keys
import pyodbc
import sqlalchemy
import time
from datetime import timedelta, date
import subprocess
import win32api,time,win32con
import win32com.client

shell = win32com.client.Dispatch("WScript.Shell")


def keyb(ch=None,shift=False,control=False,alt=False, delaik=0.02):
    for b in ch:
        c=b
        if (b>='A' and b<='Z') or shift:
            win32api.keybd_event(win32con.VK_SHIFT, 0, 0, 0)
        if b>='a' and b<='z':
            c=b.upper()
        if alt:
            win32api.keybd_event(win32con.VK_MENU, 0, 0, 0)
            time.sleep(0.250)
        if control:
            win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
        if isinstance(b,(int)):
            cord=b
        else:
            cord=ord(c)

        win32api.keybd_event(cord, 0, win32con.KEYEVENTF_EXTENDEDKEY | 0, 0)
        if delaik>0.0:
            time.sleep(delaik)
        win32api.keybd_event(cord, 0, win32con.KEYEVENTF_EXTENDEDKEY | win32con.KEYEVENTF_KEYUP, 0)
        if delaik>0.0:
            time.sleep(delaik)

        if control:
            win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
        if alt:
            win32api.keybd_event(win32con.VK_MENU, 0, win32con.KEYEVENTF_KEYUP, 0)
            time.sleep(0.05)
        if (b>='A' and b<='Z') or shift:
            win32api.keybd_event(win32con.VK_SHIFT, 0, win32con.KEYEVENTF_KEYUP, 0)


# Busca os dados da planilha para pesquisar nos sites
# desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
# arquivo = '\Certificado_Chrome.exe'
# desktop = desktop_path + arquivo





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

def diff_days(date1, date2):
    d1 = datetime.strptime(date1, "%d-%m-%Y")
    d2 = datetime.strptime(date2, "%d-%m-%Y")
    return abs((date2 - date1).days)

def frame_switch(name):
  browser.switch_to.frame(browser.find_element_by_name(name))



# Conexao com o servidor SQL
server = '52.67.126.131'
database = 'DataWareHouse'
username = 'powerbi'
password = 'powerbi2018'
cnxn = pyodbc.connect('DRIVER={ODBC Driver 13 for SQL Server};SERVER=' + server + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password)
cursor = cnxn.cursor()

# Conexao 2
engine = sqlalchemy.create_engine('mssql+pyodbc://powerbi:powerbi2018@52.67.126.131/DataWareHouse?driver=ODBC+Driver+13+for+SQL+Server')
engine.connect()


# browser = webdriver.Chrome(desired_capabilities=caps)
browser = webdriver.Chrome(executable_path=r'C:/Python/Selenium/chromedriver.exe')  # Optional argument, if not specified will search path
# browser.fullscreen_window()
browser.maximize_window()
action = ActionChains(browser)


# Data de hoje
hoje = date.today().strftime('%d/%m/%Y')
data_ini = date.today() - timedelta(30)
data_ini = data_ini.strftime('%d/%m/%Y')
hoje = hoje.replace('/','',2)
data_ini = data_ini.replace('/','',2)


# browser.get('http://app.sefaz.es.gov.br/dbsit1403_agencia_virtual/')



def loginpage():
    browser.get('http://app.sefaz.es.gov.br/dbsit1403_agencia_virtual/')
    #browser.fullscreen_window()
    browser.maximize_window()
    subprocess.Popen("C:\Python\Sefaz-ES\Certificado_Chrome.exe")
    print("Efetuando Login")
    btn_prosseguir = check_exists_by_xpath("//*[@id='btnFechar']")
    if btn_prosseguir == True:
        btn_prosseguir = browser.find_element_by_xpath("//*[@id='btnFechar']")
        browser.execute_script("arguments[0].click();", btn_prosseguir)

    # "Segurança do Windows" - Class: "Credential Dialog Xaml Host" Handle: "0x000E050E"


    janelas = browser.window_handles
    qtd_janelas = len(janelas)
    if qtd_janelas > 1:
        msg = browser.switch_to().alert().getText()
        browser.switch_to().alert().accept()

    time.sleep(2)

    btn_cert = check_exists_by_xpath("//*[@id='CEWFR_RNDR1_imgSubstituto']")
    if btn_cert == True:
        btn_cert = browser.find_element_by_xpath("//*[@id='CEWFR_RNDR1_imgSubstituto']")
        certificado = browser.window_handles
        browser.execute_script("arguments[0].click();", btn_cert)



    login = check_exists_by_name("CEWFR_RNDR1$eacNRCPFCO")
    if login == True:

        login = browser.find_element_by_name("CEWFR_RNDR1$eacNRCPFCO")
        senha = browser.find_element_by_name("CEWFR_RNDR1$eacTSENHA")
        btn_login = browser.find_element_by_name("CEWFR_RNDR1$Button1")
        login.send_keys('00031831737')
        time.sleep(1)
        senha.send_keys('1212MAIS')
        browser.execute_script("arguments[0].click();", btn_login)



    btn_reconnet = check_exists_by_name("CEWFR_RNDR1$BtnReConnect")
    if btn_reconnet == True:
        btn_reconnet = browser.find_element_by_name("CEWFR_RNDR1$COL1")
        btn_reconnet.click()
        # browser.execute_script("arguments[0].click();", btn_reconnet)
        loginpage()

loginpage()

cnpj1 = check_exists_by_name("CEWFR_RNDR1$COL1")
if cnpj1 == True:
    cnpj1 = browser.find_element_by_name("CEWFR_RNDR1$COL1")
    browser.execute_script("arguments[0].click();", cnpj1)

cons_nfs = check_exists_by_xpath("//*[@id='MenuBar1']/li[7]/ul/li[12]/a")
if cons_nfs == True:
    cons_nfs = browser.find_element_by_xpath("//*[@id='MenuBar1']/li[7]/ul/li[12]/a")
    browser.execute_script("arguments[0].click();", cons_nfs)

btn_dest = check_exists_by_xpath("//*[@id='CEWFR_RNDR1_ddlIdentificador']/option[2]")
if btn_dest == True:
    btn_dest = browser.find_element_by_xpath("//*[@id='CEWFR_RNDR1_ddlIdentificador']/option[2]")
    btn_dest.click()


field_dataini = check_exists_by_name("CEWFR_RNDR1$txtdataInicio")
if field_dataini == True:
    # field_dataini = browser.find_element_by_id("ui-datepicker-div")
    time.sleep(1)
    field_dataini = browser.find_element_by_name("CEWFR_RNDR1$txtdataInicio")
    # browser.execute_script("arguments[0].click();", field_dataini)
    field_dataini.click()
    time.sleep(2)
    field_dataini.send_keys(data_ini + hoje)
    # time.sleep(2)
    # field_fim = browser.find_element_by_name("CEWFR_RNDR1$txtdataFim")
    # browser.execute_script("arguments[0].click();", field_fim)
    # field_fim.click()
    time.sleep(1)
    # field_fim.send_keys(hoje)
    btn_pesq = browser.find_element_by_name("CEWFR_RNDR1$botaoPesquisar")
    time.sleep(2)
    btn_pesq.click()
    
alerta = check_exists_by_xpath("//*[@id='Form1']/script[7]")
erro_pag = check_exists_by_xpath("/html/body/span/h2/i")

# if alerta == True:
#     browser.close()
#     browser = webdriver.Chrome(executable_path=r'C:/Python/Selenium/chromedriver.exe')
#     loginpage()

if erro_pag == True:
    browser.close()
    browser = webdriver.Chrome(executable_path=r'C:/Python/Selenium/chromedriver.exe')
    loginpage()

qtd_reg_pag = check_exists_by_xpath("//*[@id='CEWFR_RNDR1_paginacao_NFE']/table/tbody/tr/td/span[2]")
if qtd_reg_pag == True:
    qtd_reg_pag = browser.find_element_by_xpath("//*[@id='CEWFR_RNDR1_paginacao_NFE']/table/tbody/tr/td/span[2]").text
    qtd_reg_total = browser.find_element_by_xpath("//*[@id='CEWFR_RNDR1_paginacao_NFE']/table/tbody/tr/td/span[3]").text

loop_pag = 1
tpaginas = math.ceil(int(qtd_reg_total) / int(qtd_reg_pag))

print("Limpando tabela temporária")
cursor.execute("DELETE FROM Tabela_MDE_TEMP")
cursor.commit()

while loop_pag <= tpaginas:

    html_1 = browser.find_element_by_id("CEWFR_RNDR1_gvDestinatario").get_attribute('innerHTML')
    df = pd.read_html(" <table> " + html_1 + " </table> ")
    print("Salvando dados da página: " + str(loop_pag))
    print(df)
    # Salva Dados no SQL

    df[0].columns = ['Chave_Acesso', 'CNPJ_CPF', 'NF', 'Data_Emissao', 'Situacao', 'Lupa']
    df = df[0].astype(str)

    df.to_sql(
        name='Tabela_MDE_Temp',  # database table name
        con=engine,
        if_exists='append',
        index=False,
    )


    if loop_pag <= tpaginas:
        btn_prox_pag = check_exists_by_xpath("//*[@id='CEWFR_RNDR1_paginacao_NFE']/div/span[2]")
        if btn_prox_pag == True:
            btn_prox_pag = browser.find_element_by_xpath("//*[@id='CEWFR_RNDR1_paginacao_NFE']/div/span[2]")
            btn_prox_pag.click()
            loop_pag = loop_pag + 1


# print("Limpando sujeira da tabela temporária Tabela_MDE_Temp")
# cursor.execute("DELETE FROM Tabela_MDE_Temp  WHERE [0] = 'Chave Acesso'")
# cursor.commit()
print("**** FIM ****")
engine.dispose()
cursor.close()
browser.close()
subprocess.Popen("C:\Python\Sefaz-ES\Certificado_Chrome_Fechar.exe")