
import time
import pandas as pd
import os
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium import webdriver

chrome_options = webdriver.chrome.options.Options()
chrome_options.add_argument("--headless")
# chrome_options.add_argument('--disable-gpu')
# Starts driver
driver = webdriver.Chrome("C:\chromedriver\chromedriver.exe",options=chrome_options)
# driver = webdriver.PhantomJS()
# Gets page
driver.get("https://extranet.btgpactual.com/")
# Sets wait
wait = WebDriverWait(driver, 300)

time.sleep(2)
login = "tiago.nunes"
# sends login]
driver.find_element_by_id("txtLogin").send_keys(login)
# Goes to Keyboard
driver.find_element_by_xpath("//*[@id='btnValidarLogin']").click()
time.sleep(2)
# Tries to bypass keyboard

#validates
driver.find_element_by_xpath("//*[@id='btnValidate']/span").click()
time.sleep(2)

driver.find_element_by_xpath("//*[@id='menu']/li[2]/a").click()
time.sleep(2)

driver.find_element_by_xpath("//*[@id='liCotas']/a/span").click()
time.sleep(1)

select = Select(driver.find_element_by_xpath("//*[@id='ddlFundos']"))

#565123=BALAN 550607=GRAPH
select.select_by_value("565123")

#driver.find_element_by_xpath("//*[@id='ddlIndexadores_chosen']/ul/li/input").send_keys("Cdi")
#time.sleep(1)

driver.find_element_by_xpath("//*[@id='ddlIndexadores_chosen']").click()
time.sleep(1)

driver.find_element_by_xpath("//*[@id='ddlIndexadores_chosen']/div/ul/li[1]").click()
time.sleep(1)

#Xml IBOV //*[@id="ddlIndexadores_chosen"]/div/ul/li[6]
#Xml First Date //*[@id="txtInicio"]
#Xml Second Date //*[@id="txtFim"]

driver.find_element_by_xpath("//*[@id='linkbtconsultar']/a").click()
time.sleep(2)

driver.find_element_by_xpath("//*[@id='tblAtivoCarteira_wrapper']/div[2]/div[3]/div/div[2]/a[1]").click()
time.sleep(3)

driver.close()

#result_file = os.path.join("C:\Downloads","Result.xls")

#if os.path.exists("Balandiario.xlsx"):
#        os.remove("Balandiario.xlsx")

#os.rename(result_file, "Balandiario.xls")
