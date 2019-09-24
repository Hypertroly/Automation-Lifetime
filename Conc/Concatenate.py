import numpy as np
import pandas as pd
import os
import math
import win32com.client as win32

crie = os.path.join(r"C:\Downloads","cri.xlsx")
crae = os.path.join(r"C:\Downloads","cra.xlsx")
debe = os.path.join(r"C:\Downloads","deb.xlsx")

if os.path.exists(crie):
    os.remove(crie)
    os.remove(crae)
    os.remove(debe)

cri = r"C:\Downloads\cri.xls"
cra = r"C:\Downloads\cra.xls"
deb = r"C:\Downloads\deb.xls"
excel = win32.Dispatch('Excel.Application')
wbcri = excel.Workbooks.Open(cri)
wbcra = excel.Workbooks.Open(cra)
wbdeb = excel.Workbooks.Open(deb)

wbcri.SaveAs(cri+"x", FileFormat = 51)
wbcra.SaveAs(cra+"x", FileFormat = 51)
wbdeb.SaveAs(deb+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
wbcri.Close()                               #FileFormat = 56 is for .xls extension
wbcra.Close()
wbdeb.Close()
excel.Application.Quit()

print("Arquivo convertido para xlsx")

#CRI
try:
    #Read Excel
    i = pd.read_excel(r'C:\Downloads\cri.xlsx', 'cri', index_col=None, na_values=['NA'], skiprows=8)

    #Remove columns
    i = i.drop(columns=['Núm. de Negócios','Qtde Negociada', 'Data Vencto', 'Preço Médio', 'Preço Máximo', 'Valor na Curva', 'Valor Financeiro'])

    #Gets first 11 numbers
    i['Ativo'] = i.Ativo.astype(str).str[:11].astype(str)


except IOError:
    print("Excel CRI não encontrado")

finally:
    print("CRI criado com sucesso")

#CRA
try:
    #Read Excel
    a = pd.read_excel(r'C:\Downloads\cra.xlsx', 'cra', index_col=None, na_values=['NA'], skiprows=8)

    #Remove columns
    a = a.drop(columns=['Núm. de Negócios','Qtde Negociada', 'Data Vencto', 'Preço Médio', 'Preço Máximo', 'Valor na Curva', 'Valor Financeiro'])


except IOError:
    print("Excel CRA não encontrado")

finally:
    print("CRA criado com sucesso")

#DEB
try:
    #Read Excel
    d = pd.read_excel(r'C:\Downloads\deb.xlsx', 'deb', index_col=None, na_values=['NA'], skiprows=7)

    #Remove columns
    d = d.drop(columns=['Núm. de Negócios','Qtde Negociada', 'Preço Médio', 'Preço Máximo', 'Valor Financeiro', 'Liquidação'])
    #Gets first 6 numbers
    d['Ativo'] = d.Ativo.astype(str).str[:6].astype(str)

except IOError:
    print("Excel DEB não encontrado")

finally:
    print("DEB criado com sucesso")

#Concatenate frames
framesconc = [a, i, d]
c = pd.concat(framesconc, sort=False)

if os.path.exists("Conc.xlsx"):
    os.remove("Conc.xlsx")

c.to_excel("Conc.xlsx", sheet_name='Conc')
