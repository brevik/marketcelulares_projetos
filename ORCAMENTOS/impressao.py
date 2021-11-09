#PROGRAMA PARA AUTOMAÇÃO DE IMPRESSOES DE ORÇAMENTOS
import pandas as pd
from datetime import datetime
import win32print
import win32api
import time

nomeImpressora = 'POS58 Printer(3)' #Nome da impressora que deseja imprimir (encontrado em "Impressora e scanner" no windows)
nomeImpressoraPadrao = 'PDFCreator' #Nome da impressora que deseja deixar como padrão

caminho = r"C:\Users\Vanderson\Documents\DRIVER\MARKETCELULARES"
arquivoImpressao = "IMPRESSAO_ORCAMENTOS.txt"
arquivoExcel = 'ORCAMENTOS.xlsx'

tabela = pd.read_excel(arquivoExcel, sheet_name=1)

texto = datetime.today().strftime('%d/%m/%Y')
texto += "\n============\n"
texto += "---------------\n"

for cel,q in enumerate(tabela.values):
	texto += (str(q[0]) + '\n' 
	+ str(q[1]) + '\n' 
	+ str(q[2]) + ' -- ' + str(q[3]) + '\n' 
	+ 'Senha: ' + str(q[5]) + '\n' 
	+ str(q[4]) )
	texto += "\n\n-------------\n"

print(texto)

arq = open(arquivoImpressao, "w")
arq.write(texto)
arq.close()

lista_impressoras = win32print.EnumPrinters(2) # 2 retona informações das impressoras

i = 0
idImpressora = 0
for imp in lista_impressoras:
	if imp[2] == nomeImpressora:
		idImpressora = i #pegar id da impressora a ser usada
	#print(imp)
	i+=1

impressora = lista_impressoras[idImpressora] #numero da impressora na lista
win32print.SetDefaultPrinterW(nomeImpressora) #passar o nome da impressora na lista e define como padrao
win32api.ShellExecute(0, "print", arquivoImpressao, None, caminho,0)
time.sleep(0.5)
win32print.SetDefaultPrinterW(nomeImpressoraPadrao) #definir impressora PDF como padrao novamente


	