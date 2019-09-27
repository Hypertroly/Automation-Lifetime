from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium import webdriver
import os
import sys
import datetime
import win32com.client as win32
import image_slicer
from image_slicer import join
from PIL import Image
import numpy as np


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
driver.maximize_window()

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

#Go to the desired page on 

driver.find_element_by_xpath("//*[@id='menu']/li[2]/a").click()
sleep(2)

driver.find_element_by_xpath("//*[@id='liCotas']/a/span").click()
sleep(1)

select = Select(driver.find_element_by_xpath("//*[@id='ddlFundos']"))
sleep(3)

#565123=BALAN 550607=GRAPH
select.select_by_value("565123")

driver.find_element_by_xpath("//*[@id='linkbtconsultar']/a").click()
sleep(3)

ctbl=0
ctgr=0

datahj = datetime.date.today()

datahj = str(datahj)

dia = int(datahj[8]+datahj[9])

dia = str(dia-1)

datahj = dia + '/' + datahj[5]+datahj[6] + '/' + datahj[0:4]

databalan = driver.find_element_by_xpath('//*[@id="tblAtivoCarteira"]/tbody/tr/td[2]').text

#Confirms if the day is today
if databalan==datahj:
    if os.path.exists('PrintBalan.png'):
        os.remove('PrintBalan.png')
    sleep(5)
    driver.save_screenshot('PrintBalan.png')
else:
    ctbl=1

driver.refresh()

select = Select(driver.find_element_by_xpath("//*[@id='ddlFundos']"))
sleep(3)

select.select_by_value("550607")

driver.find_element_by_xpath("//*[@id='linkbtconsultar']/a").click()
sleep(3)

#d = datetime.date(y,m,d)
datagraph = driver.find_element_by_xpath('//*[@id="tblAtivoCarteira"]/tbody/tr/td[2]').text

if datagraph==datahj:
    if os.path.exists('PrintGraph.png'):
        os.remove('PrintGraph.png')
    sleep(5)
    driver.save_screenshot('PrintGraph.png')
else:
    ctgr=1

if ctbl==1 and ctgr==1:
    print('Cotas Balanced e Graphene não atualizadas')
elif ctbl==1 and ctgr==0:
    print('Cota Balanced atualizada, cota Graphene desatualizada')
elif ctbl==0 and ctgr==1:
    print('Cota Balanced desatualizada, cota Graphene atualizada')
elif ctbl==0 and ctbl==0:
    print('Cotas Balanced e Graphene atualizadas')
driver.close()

image_slicer.slice(r'C:\Users\thiago.sousa\Desktop\pandas test\Print Cota\PrintGraph.png', 68)

sleep(3)

a = r'C:\Users\thiago.sousa\Desktop\pandas test\Print Cota\PrintGraph_08_01.png'
b = r'C:\Users\thiago.sousa\Desktop\pandas test\Print Cota\PrintGraph_08_02.png'
c = r'C:\Users\thiago.sousa\Desktop\pandas test\Print Cota\PrintGraph_08_03.png'
list_im = [a,b,c]
imgs    = [ Image.open(i) for i in list_im ]
# pick the image which is the smallest, and resize the others to match it (can be arbitrary image shape here)
min_shape = sorted( [(np.sum(i.size), i.size ) for i in imgs])[0][1]
imgs_comb = np.hstack( (np.asarray( i.resize(min_shape) ) for i in imgs ) )

# save that beautiful picture
imgs_comb = Image.fromarray( imgs_comb)
imgs_comb.save( 'FotoCotaGraph.png' )

image_slicer.slice(r'C:\Users\thiago.sousa\Desktop\pandas test\Print Cota\PrintBalan.png', 68)

sleep(3)

a = r'C:\Users\thiago.sousa\Desktop\pandas test\Print Cota\PrintBalan_08_01.png'
b = r'C:\Users\thiago.sousa\Desktop\pandas test\Print Cota\PrintBalan_08_02.png'
c = r'C:\Users\thiago.sousa\Desktop\pandas test\Print Cota\PrintBalan_08_03.png'
list_im = [a,b,c]
imgs    = [ Image.open(i) for i in list_im ]
# pick the image which is the smallest, and resize the others to match it (can be arbitrary image shape here)
min_shape = sorted( [(np.sum(i.size), i.size ) for i in imgs])[0][1]
imgs_comb = np.hstack( (np.asarray( i.resize(min_shape) ) for i in imgs ) )

# save that beautiful picture
imgs_comb = Image.fromarray( imgs_comb)
imgs_comb.save( 'FotoCotaBalan.png' )

datahj = str(datetime.date.today())

datahj = datahj[8]+ datahj[9] + '/' + datahj[5]+datahj[6] + '/' + datahj[0:4]

outlook = win32.Dispatch('Outlook.Application')

sendfromAC=None
for oacc in outlook.Session.Accounts:
    if oacc.SmtpAddress == "cota.fundo@xpi.com.br":
#    if oacc.SmtpAddress == 'Tiago Sousa <tiago.sousa@LIFETIMEASSET.COM.BR>':
        sendfromAC = oacc
        break

mail = outlook.CreateItem(0)

if sendfromAC:
    mail._oleobj_.Invoke(*(64209, 0, 8, 0, sendfromAC))
mail.To = 'cota.fundo@xpi.com.br'
#mail.To = 'Tiago Sousa <tiago.sousa@LIFETIMEASSET.COM.BR>'
mail.Cc = 'Operations <operations@LIFETIMEASSET.COM.BR>'
mail.Subject = 'Lifetime Asset - Conciliação de Movimentações'
mail.Attachments.Add(r"C:\Downloads\Results.xls")

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
                + "<BR><BR> Segue a planilha com as movimentações do dia. </b> "\
#                + "<BR><BR> Att, </b>"\
#                + "<html><body><img src='cid:Assinatura'></body></html>"

a=''
while a != 'y' and 'n':
    a=input('Email Balanced completo, deseja enviá-lo? (y/n) ')

    if a=='n':
        sys.exit()

mail.Send()
print('Email enviado')

outlook = win32.Dispatch('Outlook.Application')

sendfromAC=None
for oacc in outlook.Session.Accounts:
    if oacc.SmtpAddress == "cota.fundo@xpi.com.br":
#    if oacc.SmtpAddress == 'Tiago Sousa <tiago.sousa@LIFETIMEASSET.COM.BR>':
        sendfromAC = oacc
        break

mail = outlook.CreateItem(0)

if sendfromAC:
    mail._oleobj_.Invoke(*(64209, 0, 8, 0, sendfromAC))
mail.To = 'cota.fundo@xpi.com.br'
#mail.To = 'Tiago Sousa <tiago.sousa@LIFETIMEASSET.COM.BR>'
mail.Cc = 'Operations <operations@LIFETIMEASSET.COM.BR>'
mail.Subject = 'Lifetime Asset - Cota Diária GRAPHENE FIA'
mail.Attachments.Add(r"C:\Users\thiago.sousa\Desktop\pandas test\Print Cota\FotoCotaGraph.png")

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
                + "<BR><BR> Segue cota referente ao dia %s  </b> "% (datahj)\
#                + "<BR><BR> Att, </b>"\
#                + "<html><body><img src='cid:Assinatura'></body></html>"

a=''
while a != 'y' and 'n':
    a=input('Email Graphene completo, deseja enviá-lo? (y/n) ')

    if a=='n':
        sys.exit()

mail.Send()
print('Email enviado')