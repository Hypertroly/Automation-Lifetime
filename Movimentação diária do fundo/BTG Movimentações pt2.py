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
driver.find_element_by_xpath("//*[@id='contentVirtualKeyboard']/div/div/div[2]").click()
driver.find_element_by_xpath("//*[@id='contentVirtualKeyboard']/div/div/div[11]").click()
driver.find_element_by_xpath("//*[@id='contentVirtualKeyboard']/div/div/div[4]").click()
driver.find_element_by_xpath("//*[@id='contentVirtualKeyboard']/div/div/div[54]").click()
driver.find_element_by_xpath("//*[@id='contentVirtualKeyboard']/div/div/div[18]").click()
driver.find_element_by_xpath("//*[@id='contentVirtualKeyboard']/div/div/div[21]").click()
driver.find_element_by_xpath("//*[@id='contentVirtualKeyboard']/div/div/div[27]").click()
driver.find_element_by_xpath("//*[@id='contentVirtualKeyboard']/div/div/div[31]").click()
driver.find_element_by_xpath("//*[@id='contentVirtualKeyboard']/div/div/div[22]").click()
#validates
driver.find_element_by_xpath("//*[@id='btnValidate']/span").click()
sleep(2)

#Go to the desired page on BTG
driver.find_element_by_xpath('//*[@id="menu"]/li[3]/a').click()

driver.find_element_by_xpath('//*[@id="liMovimentacoes"]/a/span').click()

driver.find_element_by_xpath('//*[@id="linkbtConsultarCL"]/a/span').click()
sleep(3)

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
            driver.close()
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

#Changes format of excel to send to XP
if os.path.exists(r"C:\Downloads\Lifetime.txt"):
    os.remove(r"C:\Downloads\Lifetime.txt")

try:
    fname = r"C:\Downloads\Results.xls"
    excel = win32.Dispatch('Excel.Application')
    wb = excel.Workbooks.Open(fname)
    filename = r"C:\Downloads\Lifetime.txt"

    wb.SaveAs(filename, FileFormat = 42)    #FileFormat = 51 is for .xlsx extension
    wb.Close()                               #FileFormat = 56 is for .xls extension
    excel.Application.Quit()

    print("Arquivo convertido para txt")
except AttributeError:
    print("Feche o processo EXCEL no gerenciador de tarefas e execute o programa novamente")

outlook = win32.Dispatch('Outlook.Application')

sendfromAC=None
for oacc in outlook.Session.Accounts:
    if oacc.SmtpAddress == "Movimentacoes Fundos - XP Investimentos <movimentacoes.fundos@xpi.com.br>":
#    if oacc.SmtpAddress == 'Tiago Sousa <tiago.sousa@LIFETIMEASSET.COM.BR>':
        sendfromAC = oacc
        break

mail = outlook.CreateItem(0)

if sendfromAC:
    mail._oleobj_.Invoke(*(64209, 0, 8, 0, sendfromAC))
mail.To = 'Movimentacoes Fundos - XP Investimentos <movimentacoes.fundos@xpi.com.br>'
#mail.To = 'Tiago Sousa <tiago.sousa@LIFETIMEASSET.COM.BR>'
mail.Cc = 'Operations <operations@LIFETIMEASSET.COM.BR>'
mail.Subject = 'XP Movimentações - Lifetime'
mail.Attachments.Add(r"C:\Downloads\Lifetime.txt")

#attachment1 = mail.Attachments.Add(r'C:\movbtg\Assinatura.png')
#attachment1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "Assinatura")

#mail.HTMLBody = "<HTML lang='en' xmlns='http://www.w3.org/1999/xhtml' xmlns:o='urn:schemas-microsoft-com:office:office'> " \
#                + "<head>" \
#                + "<!--[if gte mso 9]><xml> \
#                        <o:OfficeDocumentSettings> \
#                        <o:Allowjpeg/> \
#                        <o:PixelsPerInch>96</o:PixelsPerInch> \
#                        </o:OfficeDocumentSettings> \
#                    </xml> \
#                    <![endif]-->" \
#                + "</head>" \
#                + "<BODY>"



mail.HTMLBody = mail.HTMLBody + "<BR>XP,<b> </b>" \
                + "<BR><BR> Segue o arquivo .txt com as movimentações do dia. </b> "\
#                + "<BR><BR> Att, </b>"\
#                + "<html><body><img src='cid:Assinatura'></body></html>"

a=''
while a != 'y' and 'n':
    a=input('Email completo, deseja enviá-lo? (y/n) ')

    if a=='n':
        sys.exit()

mail.Send()
print('Email enviado')