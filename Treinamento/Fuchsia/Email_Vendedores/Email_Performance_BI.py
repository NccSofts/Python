import os
import pandas as pd
import pyodbc
import sqlalchemy
from openpyxl import load_workbook
import re
import time
from datetime import timedelta, date
import PIL
from PIL import Image
import codecs
import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
import datetime


# Data de hoje
hoje = date.today().strftime('%d/%m/%Y')
diasemana = datetime.datetime.today().weekday()


# Definir conexao 0 = PowerBiV2 / 1 = Datawarehouse
opcao = 0

configuracoes = pd.read_csv("C:/Python/database.cfg", header = 0, delimiter=";")

server = configuracoes['IP'].values[opcao]
database = configuracoes['Banco'].values[opcao]
username = configuracoes['Usuario'].values[opcao]
password = configuracoes['Senha'].values[opcao]


sql_string = 'mssql+pyodbc://' + username + ':' + password + '@' + server + '/' + database + '?driver=ODBC+Driver+13+for+SQL+Server'

cnxn = pyodbc.connect(
    'DRIVER={ODBC Driver 13 for SQL Server};SERVER=' + server + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password)
cursor = cnxn.cursor()

# Conexao com o servidor SQL
engine = sqlalchemy.create_engine(sql_string)
engine.connect()






df = pd.read_sql_table('Carteira_Clientes_Resumo_Email', engine)
df.fillna(0)

lin = 0


if diasemana < 5:
    for row in df.itertuples():
        nome_vendedor = str(df['Nome'].values[lin])
        meta_formatada = df['MetaFormatada'].values[lin]
        venda_formatada = df['VendaFormatada'].values[lin]
        perc_ativos = "{0:.0%}".format(df['P_Cont_Ativos'].values[lin])
        perc_inativos = "{0:.0%}".format(df['P_Cont_Inativos'].values[lin])
        perc_meta = "{0:.0%}".format(df['P_Meta'].values[lin])
        perc_contatos = "{0:.0%}".format(df['P_Contatos'].values[lin])
        perc_conv_contatos = "{0:.00%}".format(df['P_Conv'].values[lin])
        qtd_contatos = "{:.0f}".format(df['QtdContEfetivos'].values[lin])
        qtd_carteira = "{:.0f}".format(df['GrupoCarteira'].values[lin])
        mes_atual = df['Mes_Atual'].values[lin]
        qtd_lig_efetivas = "{:.0f}".format(df['QtdLigEfetivas'].values[lin])
        qtd_lig_outros = "{:.0f}".format(df['QtdLigOutros'].values[lin])
        qtd_lig_tentativas = "{:.0f}".format(df['QtdLigTentativas'].values[lin])
        qtd_cont_efetivas = "{:.0f}".format(df['QtdContEfetivos'].values[lin])
        qtd_cont_outros = "{:.0f}".format(df['QtdContOutros'].values[lin])
        qtd_cont_tentativas = "{:.0f}".format(df['QtdContTentativas'].values[lin])
        email_vendedor = str(df['Email'].values[lin])
        ticket_medio = str(df['Ticket_Medio_Formatado'].values[lin])

        # Criando Grafico Meta Venda
        percent_meta = int(df['P_Meta'].values[lin] * 100)  # Percent for gauge
        output_file_name = 'C:/Python/Gauge/Meta.png'

        # X and Y coordinates of the center bottom of the needle starting from the top left corner
        #   of the image
        x = 825
        y = 825
        loc = (x, y)

        percent = percent_meta / 100
        rotation = 180 * percent  # 180 degrees because the gauge is half a circle
        rotation = 90 - rotation  # Factor in the needle graphic pointing to 50 (90 degrees)

        dial = Image.open('C:/Python/Gauge/needle.png')
        dial = dial.rotate(rotation, resample=PIL.Image.BICUBIC, center=loc)  # Rotate needle

        gauge = Image.open('C:/Python/Gauge/gauge.png')
        gauge.paste(dial, mask=dial)  # Paste needle onto gauge
        gauge.save(output_file_name)

        # Criando Grafico Qtd Contatos
        percent_contato = int(df['P_Contatos'].values[lin] * 100)  # Percent for gauge
        output_file_name = 'C:/Python/Gauge/Contatos.png'

        # X and Y coordinates of the center bottom of the needle starting from the top left corner
        #   of the image
        x = 825
        y = 825
        loc = (x, y)

        percent = percent_contato / 100
        rotation = 180 * percent  # 180 degrees because the gauge is half a circle
        rotation = 90 - rotation  # Factor in the needle graphic pointing to 50 (90 degrees)

        dial = Image.open('C:/Python/Gauge/needle.png')
        dial = dial.rotate(rotation, resample=PIL.Image.BICUBIC, center=loc)  # Rotate needle

        gauge = Image.open('C:/Python/Gauge/gauge.png')
        gauge.paste(dial, mask=dial)  # Paste needle onto gauge
        gauge.save(output_file_name)



        # Criando Grafico CLientes Ativos
        percent_contato = int(df['P_Cont_Ativos'].values[lin] * 100)  # Percent for gauge
        output_file_name = 'C:/Python/Gauge/Ativos.png'

        # X and Y coordinates of the center bottom of the needle starting from the top left corner
        #   of the image
        x = 825
        y = 825
        loc = (x, y)

        percent = percent_contato / 100
        rotation = 180 * percent  # 180 degrees because the gauge is half a circle
        rotation = 90 - rotation  # Factor in the needle graphic pointing to 50 (90 degrees)

        dial = Image.open('C:/Python/Gauge/needle.png')
        dial = dial.rotate(rotation, resample=PIL.Image.BICUBIC, center=loc)  # Rotate needle

        gauge = Image.open('C:/Python/Gauge/gauge.png')
        gauge.paste(dial, mask=dial)  # Paste needle onto gauge
        gauge.save(output_file_name)



        # Criando Grafico CLientes Inativos
        percent_contato = int(df['P_Cont_Inativos'].values[lin] * 100)  # Percent for gauge
        output_file_name = 'C:/Python/Gauge/Inativos.png'

        # X and Y coordinates of the center bottom of the needle starting from the top left corner
        #   of the image
        x = 825
        y = 825
        loc = (x, y)

        percent = percent_contato / 100
        rotation = 180 * percent  # 180 degrees because the gauge is half a circle
        rotation = 90 - rotation  # Factor in the needle graphic pointing to 50 (90 degrees)

        dial = Image.open('C:/Python/Gauge/needle.png')
        dial = dial.rotate(rotation, resample=PIL.Image.BICUBIC, center=loc)  # Rotate needle

        gauge = Image.open('C:/Python/Gauge/gauge.png')
        gauge.paste(dial, mask=dial)  # Paste needle onto gauge
        gauge.save(output_file_name)



        # Criando HTML para o corpo do Email
        nome = df['Nome'].values[lin]
        linha_vazia = '<p>&nbsp;</p>'

        cabecalho = '<h1 style="text-align: center;"><strong>PERFORMANCE DE VENDAS - '+ mes_atual +'&nbsp;</strong></h1>'
        html_lin1 = '<h2 style="text-align: center;"><p><strong>Vendedor: ' + str(df['Nome'].values[lin]) + '</strong></p><h2>'
        html_lin2 = '<p style="text-align: center;"><strong>Meta de Venda: ' + str(meta_formatada) + '</strong></p>'
        html_lin3 = '<p style="text-align: center;"><strong>Venda Realizada: ' + str(venda_formatada) + '</strong></p>'
        html_lin4 = '<p style="text-align: center;"><strong>Ticket Médio: ' + str(ticket_medio) + '</strong></p>'
        html_contatos = '<p style="text-align: center;"><strong>Qtd Contatos: ' + str(qtd_contatos) + ' / Carteira: '+ qtd_carteira +'</strong></p>'





        # Cabeçalho do Email
        with open("C:\Python\Gauge\corpoemail.html", "w") as file:
            file.write(linha_vazia)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(cabecalho)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(html_lin1)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(html_lin2)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(html_lin3)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(html_lin4)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(html_contatos)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(linha_vazia)

        # Imagens do email
        tab_img1 = '<table style="width: 600px; margin-left: auto; margin-right: auto;">'
        tab_img2 = '<tbody>'
        tab_img3 = '<tr>'
        tab_img4 = '<td style="width: 250px; text-align: center;">' + perc_meta + ' (Meta Venda)</td>'
        tab_img5 = '<td style="width: 37.5px; text-align: center;">&nbsp;</td>'
        tab_img6 = '<td style="width: 250px; text-align: center;">' + perc_contatos + ' (% Contatos)</td>'
        tab_img7 = '</tr>'
        tab_img8 = '<tr>'
        tab_img9 = '<td style="width: 250px;"><img style="display: block; margin-left: auto; margin-right: auto;" src="cid:image2" alt="" width="250" height="130" /></td>'
        tab_img10 = '<td style="width: 37.5px;">&nbsp;</td>'
        tab_img11 = '<td style="width: 250px;"><img style="display: block; margin-left: auto; margin-right: auto;" src="cid:image1" alt="" width="250" height="130" /></td>'
        tab_img12 = '</tr>'
        tab_img13 = '</tbody>'
        tab_img14 = '</table>'

        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_img1)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_img2)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_img3)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_img4)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_img5)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_img6)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_img7)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_img8)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_img9)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_img10)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_img11)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_img12)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_img13)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_img14)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(linha_vazia)


        # Imagens do email Clientes Ativos e Inativos
        tab_img1 = '<table style="width: 600px; margin-left: auto; margin-right: auto;">'
        tab_img2 = '<tbody>'
        tab_img3 = '<tr>'
        tab_img4 = '<td style="width: 250px; text-align: center;">' + perc_ativos + ' (Conversão Ativos)</td>'
        tab_img5 = '<td style="width: 37.5px; text-align: center;">&nbsp;</td>'
        tab_img6 = '<td style="width: 250px; text-align: center;">' + perc_inativos + ' (Conversão Inativos)</td>'
        tab_img7 = '</tr>'
        tab_img8 = '<tr>'
        tab_img9 = '<td style="width: 250px;"><img style="display: block; margin-left: auto; margin-right: auto;" src="cid:image3" alt="" width="250" height="130" /></td>'
        tab_img10 = '<td style="width: 37.5px;">&nbsp;</td>'
        tab_img11 = '<td style="width: 250px;"><img style="display: block; margin-left: auto; margin-right: auto;" src="cid:image4" alt="" width="250" height="130" /></td>'
        tab_img12 = '</tr>'
        tab_img13 = '</tbody>'
        tab_img14 = '</table>'

        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_img1)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_img2)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_img3)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_img4)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_img5)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_img6)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_img7)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_img8)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_img9)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_img10)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_img11)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_img12)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_img13)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_img14)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(linha_vazia)



        # TABELA LIGAÇOES
        tab_lig1 = '<table style="margin-left: auto; margin-right: auto; background-color: 02046e;" border="1">'
        tab_lig2 = '<tbody>'
        tab_lig3 = '<tr>'
        tab_lig4 = '<td>Cont. Efetivos</td>'
        tab_lig5 = '<td>Cont. Outros</td>'
        tab_lig6 = '<td>Cont. Tentativas</td>'
        tab_lig7 = '<td>Lig. Efetivas</td>'
        tab_lig8 = '<td>Lig. Outros</td>'
        tab_lig9 = '<td>Lig. Tentativas</td>'
        tab_lig10 = '</tr>'
        tab_lig11 = '<tr>'
        tab_lig12 = '<td style="text-align: center;">'+ qtd_cont_efetivas +'</td>'
        tab_lig13 = '<td style="text-align: center;">'+ qtd_cont_outros +'</td>'
        tab_lig14 = '<td style="text-align: center;">'+ qtd_cont_tentativas +'</td>'
        tab_lig15 = '<td style="text-align: center;">'+ qtd_lig_efetivas +'</td>'
        tab_lig16 = '<td style="text-align: center;">'+ qtd_lig_outros +'</td>'
        tab_lig17 = '<td style="text-align: center;">'+ qtd_lig_tentativas +'</td>'
        tab_lig18 = '</tr>'
        tab_lig19 = '</tbody>'
        tab_lig20 = '</table>'

        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_lig1)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_lig2)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_lig3)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_lig4)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_lig5)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_lig6)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_lig7)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_lig8)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_lig9)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_lig10)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_lig11)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_lig12)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_lig13)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_lig14)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_lig15)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_lig16)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_lig17)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_lig18)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_lig19)
        with open("C:\Python\Gauge\corpoemail.html", "a") as file:
            file.write(tab_lig20)


        corpo = codecs.open("C:\Python\Gauge\corpoemail.html", 'r')
        html = corpo.read()





        # ENVIA EMAIL
        diasemana = date.today().weekday()
        data_atual = date.today()
        data_pt_br = data_atual.strftime('%d/%m/%Y')
        # sender_email = "credito@maisproxima.com.br"
        user_login = 'bi@maisproxima.com'
        # receiver_email = "anderson.souza@bioscan.com.br"
        receiver_email = [email_vendedor]
        # receiver_email = ['simone.dario@maisproxima.com.br']
        # receiver_email = ['anderson.souza@maisproxima.com.br']
        # receiver_email = ['Thais Sousa <thais.sousa@bioscan.com.br>']
        # password = input("Type your password and press enter:")
        # password = 'MP#qwert1234'
        password = '+Proxima2019'

        titulo_email = 'PERFORMANCE DE VENDAS - ' + data_pt_br + ' - '+ nome_vendedor

        sender_email = 'Gestão da Informação - Mais Próxima <bi@maisproxima.com>'
        # message = MIMEMultipart("alternative")
        message = MIMEMultipart("Related")
        message["Subject"] = titulo_email
        message["From"] = sender_email
        message["To"] = ", ".join(receiver_email)

        # Turn these into plain/html MIMEText objects
        # part1 = MIMEText(text, "plain")
        part2 = MIMEText(html, "html")

        # Add HTML/plain-text parts to MIMEMultipart message
        # The email client will try to render the last part first
        # message.attach(part1)
        message.attach(part2)

        # Imagem Contatos
        # This example assumes the image is in the current directory
        fp = open('C:/Python/Gauge/Contatos.png', 'rb')
        msgImage = MIMEImage(fp.read())
        fp.close()

        # Define the image's ID as referenced above
        msgImage.add_header('Content-ID', '<image1>')
        message.attach(msgImage)

        # Imagem Meta
        fp = open('C:/Python/Gauge/Meta.png', 'rb')
        msgImage = MIMEImage(fp.read())
        fp.close()

        msgImage.add_header('Content-ID', '<image2>')
        message.attach(msgImage)


        # Imagem Clientes Ativos
        fp = open('C:/Python/Gauge/Ativos.png', 'rb')
        msgImage = MIMEImage(fp.read())
        fp.close()

        msgImage.add_header('Content-ID', '<image3>')
        message.attach(msgImage)


        # Imagem Clientes Inativos
        fp = open('C:/Python/Gauge/Inativos.png', 'rb')
        msgImage = MIMEImage(fp.read())
        fp.close()

        msgImage.add_header('Content-ID', '<image4>')
        message.attach(msgImage)


        # mailserver = smtplib.SMTP('smtp.office365.com', 587)
        mailserver = smtplib.SMTP('smtp.gmail.com', 587)
        mailserver.ehlo()
        mailserver.starttls()
        mailserver.login(user_login, password)
        mailserver.sendmail(sender_email, receiver_email, message.as_string())
        mailserver.quit()

        print('Email enviado para o vendedor '+ nome + ' ('+email_vendedor+')')
        lin = lin + 1


else:
    print('Email enviado somente nos dias uteis...')

print('Emails enviados com sucesso...')

