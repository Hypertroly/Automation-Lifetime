from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium import webdriver
import glob
import os
import sys

for name in glob.glob(r'C:\movbtg\02739_'):
    balan=name
for name in glob.glob(r'C:\movbtg\03090_'):
    graph=name

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

driver.find_element_by_xpath('//*[@id="ddlOptMovimentacao"]').click()

driver.find_element_by_xpath('//*[@id="ddlOptMovimentacao"]/option[3]').click()

driver.find_element_by_xpath('//*[@id="ddlModelo"]').click()

driver.find_element_by_xpath('//*[@id="ddlModelo"]/option[2]').click()
sleep(4)

#Send mov archives
driver.find_element_by_xpath('//*[@id="inputs-forms-arquivofundos"]/div[10]/div/div/input').sendkeys(balan)

driver.find_element_by_xpath('//*[@id="divImportarFundos"]/a/span').click()
sleep(4)

driver.find_element_by_xpath('//*[@id="spnMessageImportFundos"]/a[1]').click()
sleep(4)

driver.find_element_by_xpath('//*[@id="inputs-forms-arquivofundos"]/div[10]/div/div/input').sendkeys(graph)

driver.find_element_by_xpath('//*[@id="divImportarFundos"]/a/span').click()
sleep(4)

driver.find_element_by_xpath('//*[@id="spnMessageImportFundos"]/a[1]').click()
sleep(4)

driver.close()