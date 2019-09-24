import os
import pandas as pd
import win32com.client as win32

result_file = os.path.join("C:\Downloads","Results.xls")
new_file = os.path.join("C:\Downloads","Balandiario.xls")

if os.path.exists(new_file):
    os.remove(new_file)
if os.path.exists("C:\Downloads\Balandiario.xlsx"):
    os.remove("C:\Downloads\Balandiario.xlsx")

os.rename(result_file, new_file)

fname = "C:\Downloads\Balandiario.xls"
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(fname)

wb.SaveAs(fname+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
wb.Close()                               #FileFormat = 56 is for .xls extension
excel.Application.Quit()

try:
    new = pd.read_excel('C:\Downloads\Balandiario.xlsx', 'Balandiario', index_col=None, na_values=['NA'])

except IOError:
    print("Excel novo não encontrado")
try:
    fund = pd.read_excel(r'\\192.168.1.5\lftm_asset\GESTAO\FUNDOS ABERTOS\LIFETIME BALANCED\Lâminas\Lamina_Balanced.xlsx', 'Historico', index_col=None, na_values=['NA'])

except IOError:
   print("Excel lâmina não encontrado")

copy = new[new.columns[0]]
print(copy)
