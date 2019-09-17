import numpy as np
import pandas as pd
import os
import math

#CRI
try:
    #Read Excel
    i = pd.read_excel('CRI.xlsx', 'CRI', index_col=None, na_values=['NA'])

    #Remove columns
    i = i.drop(columns=['Núm. de Negócios','Qtde Negociada', 'Data Vencto', 'Preço Médio', 'Preço Máximo', 'Valor na Curva', 'Valor Financeiro'])

    #Gets first 11 numbers
    i['Ativo'] = i.Ativo.astype(str).str[:11].astype(str)


except Exception(FileNotFoundError):
    print("Excel CRI não encontrado")

finally:
    print("CRI criado com sucesso")

#CRA
try:
    #Read Excel
    a = pd.read_excel('CRA.xlsx', 'CRA', index_col=None, na_values=['NA'])

    #Remove columns
    a = a.drop(columns=['Núm. de Negócios','Qtde Negociada', 'Data Vencto', 'Preço Médio', 'Preço Máximo', 'Valor na Curva', 'Valor Financeiro'])


except Exception(FileNotFoundError):
    print("Excel CRA não encontrado")

finally:
    print("CRA criado com sucesso")

#DEB
try:
    #Read Excel
    d = pd.read_excel('DEB.xlsx', 'DEB', index_col=None, na_values=['NA'])

    #Remove columns
    d = d.drop(columns=['Núm. de Negócios','Qtde Negociada', 'Preço Médio', 'Preço Máximo', 'Valor Financeiro', 'Liquidação'])
    #Gets first 6 numbers
    d['Ativo'] = d.Ativo.astype(str).str[:6].astype(str)

except Exception(FileNotFoundError):
    print("Excel DEB não encontrado")

finally:
    print("DEB criado com sucesso")

#Concatenate frames
framesconc = [a, i, d]
c = pd.concat(framesconc, sort=False)

if os.path.exists("Conc.xlsx"):
    os.remove("Conc.xlsx")
    
c.to_excel("Conc.xlsx", sheet_name='Conc')