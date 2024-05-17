import glob

print("Hello World!")

a = 1
b = 2
c = a + b
print(c)



diretorio = 'C:\Mais Próxima\Itau400Retorno'

padrao = '*.ret'  # exemplo de padrão para buscar arquivos de texto

lista_arquivos = glob.glob(os.path.join(diretorio, padrao))

for arquivo in lista_arquivos:
    print(arquivo)
