from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium import webdriver
import os
import sys
import win32com.client as win32
from datetime import date

chrome_options = webdriver.chrome.options.Options()
#chrome_options.add_argument("--headless")
# chrome_options.add_argument('--disable-gpu')
# Starts driver
driver = webdriver.Chrome("C:\chromedriver\chromedriver.exe",options=chrome_options)
# driver = webdriver.PhantomJS()
# Gets page
driver.get("https://extranet.btgpactual.com/")
# Sets wait
wait = WebDriverWait(driver, 300)

sleep(2)
login = "tiago.nunes"
# sends login
driver.find_element_by_id("txtLogin").send_keys(login)
# Goes to Keyboard
driver.find_element_by_xpath("//*[@id='btnValidarLogin']").click()
sleep(2)
# Tries to bypass keyboard
driver.find_element_by_xpath("//*[@id='contentVirtualKeyboard']/div/div/div[11]").click()

#validates
driver.find_element_by_xpath("//*[@id='btnValidate']/span").click()
sleep(2)

#Go to the desired page on BTG
driver.find_element_by_xpath('//*[@id="menu"]/li[3]/a').click()

driver.find_element_by_xpath('//*[@id="liMovimentacoes"]/a/span').click()

driver.find_element_by_xpath('//*[@id="linkbtConsultarCL"]/a/span').click()
sleep(6)

#Pass through all lines, analysing if any of them are still not completed, if they are completed, generate excel
#if they still aren't, close the program
ele = driver.find_element_by_xpath('//*[@id="tblLancamentos"]/tbody/tr[1]/td[3]/span[1]')
ele = ele.get_attribute('id')
base = '//*[@id="spnStatusRegistro_'
base2 = '"]'
base3 = '//*[@id="'
#prox = '//[@id="spnStatusRegistro_2430002"]'
prox = base3+ele+base2


while 1<2:
    if driver.find_elements_by_xpath(prox):
        if driver.find_element_by_xpath(prox).text=="AGUARDANDO ABERTURA DE CONTA":
            print("Ainda há movimentações não completas")
            input("Aperte enter para fechar o programa")
            sys.exit()
        number=prox[27:-2]
        number = int(number)
        number = number+1
        number = str(number)
        prox = (base+number+base2)
    else:
        print("Todas as movimentações estão completas")
        break

if os.path.exists(r"C:\Downloads\Results.xls"):
    os.remove(r"C:\Downloads\Results.xls")

driver.find_element_by_xpath('//*[@id="tblLancamentos_wrapper"]/div[3]/div[3]/div/a[1]').click()
sleep(5)
driver.close()

print("Arquivo baixado")

#Changes format of excel for sending to XP
try:
    fname = r"C:\Downloads\Results.xls"
    excel = win32.Dispatch('Excel.Application')
    wb = excel.Workbooks.Open(fname)

    wb.SaveAs(fname+"x", FileFormat = 56)    #FileFormat = 51 is for .xlsx extension
    wb.Close()                               #FileFormat = 56 is for .xls extension
    excel.Application.Quit()

    print("Arquivo convertido para xls")
except AttributeError:
    print("Feche o processo EXCEL no gerenciador de tarefas e execute o programa novamente")

outlook = win32.Dispatch('Outlook.Application')

sendfromAC=None
for oacc in outlook.Session.Accounts:
#    if oacc.SmtpAddress == "Movimentacoes Fundos - XP Investimentos <movimentacoes.fundos@xpi.com.br>":
    if oacc.SmtpAddress == 'Kevin Freundt <kevin.freundt@LIFETIMEASSET.COM.BR>':
        sendfromAC = oacc
        break

mail = outlook.CreateItem(0)

if sendfromAC:
    mail._oleobj_.Invoke(*(64209, 0, 8, 0, sendfromAC))
#mail.To = 'Movimentacoes Fundos - XP Investimentos <movimentacoes.fundos@xpi.com.br>'
mail.To = 'Kevin Freundt <kevin.freundt@LIFETIMEASSET.COM.BR>'
#mail.Cc = 'Operations <operations@LIFETIMEASSET.COM.BR>'
mail.Subject = 'Lifetime Asset - Conciliação de Movimentações'
mail.Attachments.Add(r"C:\Downloads\Results.xls")

mail.HTMLBody = mail.HTMLBody + "<BR>XP,<b> </b>" \
                + "<BR><BR> Segue a planilha com as movimentações do dia. </b> "\
                + "<BR><BR> Att, </b>"
mail.Send()