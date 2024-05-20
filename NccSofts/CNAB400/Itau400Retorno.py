import os 
import pyodbc #sql
import glob #diretorios e arquivos
import shutil #diretorios e arquivos
import pandas as pd
from datetime import datetime #tratamento de data e hora
import parametros

server = '52.67.126.131' 
database = 'PJKTB2' 
username = 'devops' 
password = 'pe7IhicAYlW&_igi?aC7' 
cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()



from ftplib import FTP

# Conectar ao servidor FTP
ftp = FTP('ftp.example.com')
ftp.login(user='username', passwd='password')

# Listar arquivos no diretório atual
files = ftp.nlst()
print('Arquivos no diretório:', files)

# Diretório de destino no servidor FTP para onde os arquivos serão movidos
dest_dir = 'processed'  # Substitua pelo caminho do diretório de destino

# Certifique-se de que o diretório de destino existe, caso contrário, crie-o
if dest_dir not in ftp.nlst():
    ftp.mkd(dest_dir)

# Processar cada arquivo
for filename in files:
    local_filename = f'local_{filename}'  # Nome do arquivo local temporário

    # Baixar o arquivo
    with open(local_filename, 'wb') as local_file:
        ftp.retrbinary(f'RETR {filename}', local_file.write)
    print(f'Arquivo {filename} baixado com sucesso.')

    # Mover o arquivo no servidor FTP
    ftp.rename(filename, f'{dest_dir}/{filename}')
    print(f'Arquivo {filename} movido para {dest_dir}/{filename}.')

    # Opcional: deletar o arquivo local após processar
    import os
    os.remove(local_filename)
    print(f'Arquivo local {local_filename} deletado.')

# Encerrar a conexão
ftp.quit()



#ArquivoDOS = 'C:\Mais Próxima\Itau400Retorno\cob_341_400_260145_240405_00000.ret'
#ArquivoDOS = 'C:\Mais Próxima\Itau400Retorno\cob_341_400_260145_240416_00000.ret'

diretorio = 'C:\Mais Próxima\Itau400Retorno'

diretoriodestino = 'C:\Mais Próxima\Itau400Retorno\Processados'

padrao = '*.ret'  # exemplo de padrão para buscar arquivos de texto

lista_arquivos = glob.glob(os.path.join(diretorio, padrao))

print("Inico")

for ArquivoDOS in lista_arquivos:
    
  print(ArquivoDOS)
  arquivo = open(ArquivoDOS,'r')
  for linha in arquivo:
      linha = linha.rstrip()

      print(linha[394:400])
      
      if linha[0] == "1":
        
        Registro = linha[0]
        TipoInscricao =  linha[1:3]
        CNPJCPF = linha[3:17]
        Agencia = linha[17:21]
        ContaCorrente = linha[23:28]
        DAC = linha[28:29]
        UsoEmpresa = linha[37:62]
        NossoNumero = linha[62:70]
        Carteira = linha[82:85]
        NossoNumeroN3 = linha[85:93]
        DACNossoNumero = linha[93:94]
        CodigoCarteira = linha[107:108]
        Ocorrencia = linha[108:110]
        
        if linha[110:116] != ("000000" or "      "):  
          DataOcorrenciaString =  '20' + linha[114:116] + '-'  +  linha[112:114]  +  '-' + linha[110:112] 
          DataOcorrencia = datetime.strptime(DataOcorrenciaString, "%Y-%m-%d")
        else:
          DataOcorrencia = ""
        
        Docuemnto = linha[116:126]
        NNConfirmado = linha[126:134]
        
        #print(linha[146:152])
        if linha[146:152] != ("000000" or "      "):  
          VencimentoString =  '20' + linha[150:152] + '-'  +  linha[148:150]  +  '-' + linha[146:148] 
          Vencimento =  datetime.strptime(VencimentoString, "%Y-%m-%d")
        else:
          Vencimento = ""          
        
        Valor = (int(linha[152:165])/100)
        Banco = linha[165:168]
        AgenciaCobradora = linha[168:172]
        DACAgenciaCobradora = linha[172:173]
        Especie = linha[173:175]
        Tarifa = (int(linha[175:188])/100)
        IOF = (int(linha[214:227])/100)
        Abatimento = (int(linha[227:240])/100)
        Descontos = (int(linha[240:253])/100)
        Principal = (int(linha[253:266])/100)
        Juros = (int(linha[266:279])/100)
        OutrosCreditos = (int(linha[279:292])/100)
        DDA = linha[292:293]

        if linha[295:301] !=  "      ":  
          DataCreditoString =  '20' + linha[299:301] + '-'  +  linha[297:299]  +  '-' + linha[295:297] 
          DataCredito = datetime.strptime(DataCreditoString, "%Y-%m-%d")      
        else:
          DataCredito = ""

        InstrucaoCancelada  = linha[301:305]
        Pagador = linha[324:354]
        Erros  = linha[377:385]
        CodigoLiquidacao  = linha[392:394]
        Sequencial  = linha[394:400]

        comando = f"""INSERT INTO PJKTB2..Itau400Retorno
                          (
                              Empresa,
                              Registro,
                              TipoInscricao,
                              CNPJCPF,
                              Agencia,
                              ContaCorrente,
                              DAC,
                              UsoEmpresa,
                              NossoNumero,
                              Carteira,
                              NossoNumeroN3,
                              DACNossoNumero,
                              CodigoCarteira,
                              Ocorrencia,
                              DataOcorrencia,
                              Docuemnto,
                              NNConfirmado,
                              Vencimento,
                              Valor,
                              Banco,
                              AgenciaCobradora,
                              DACAgenciaCobradora,
                              Especie,
                              Tarifa,
                              IOF,
                              Abatimento,
                              Descontos,
                              Principal,
                              Juros,
                              OutrosCreditos,
                              DDA,
                              DataCredito,
                              InstrucaoCancelada,
                              Pagador,
                              Erros,
                              CodigoLiquidacao,
                              Sequencial,
                              ArquivoRetorno
                          )
                      VALUES 
                          (
                              '',
                              '{Registro}',
                              '{TipoInscricao}',
                              '{CNPJCPF}',
                              '{Agencia}',
                              '{ContaCorrente}',
                              '{DAC}',
                              '{UsoEmpresa}',
                              '{NossoNumero}',
                              '{Carteira}',
                              '{NossoNumeroN3}',
                              '{DACNossoNumero}',
                              '{CodigoCarteira}',
                              '{Ocorrencia}',
                              '{DataOcorrencia}',
                              '{Docuemnto}',
                              '{NNConfirmado}',
                              '{Vencimento}',
                              {Valor},
                              '{Banco}',
                              '{AgenciaCobradora}',
                              '{DACAgenciaCobradora}',
                              '{Especie}',
                              {Tarifa},
                              {IOF},
                              {Abatimento},
                              {Descontos},
                              {Principal},
                              {Juros},
                              {OutrosCreditos},
                              '{DDA}',
                              '{DataCredito}',
                              '{InstrucaoCancelada}',
                              '{Pagador}',
                              '{Erros}',
                              '{CodigoLiquidacao}',
                              '{Sequencial}',
                              '{ArquivoDOS}'
                          )"""      


        cursor.execute(comando)
        cursor.commit()

  arquivo.close()

  # Movendo e renomeando com os.rename
  shutil.move(ArquivoDOS, diretoriodestino)

print("Fim")




