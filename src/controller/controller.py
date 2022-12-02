from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.chrome import ChromeDriverManager
import time as time
import pyautogui as p
import pandas as pd
import xlrd
from src.model.gerando_dados_planilhas import lendo_excel
# from src.model.preenchendo_colunas import colocando_nome



#Abrindo o navegador
driver = webdriver.Chrome(ChromeDriverManager().install()) 
browser = driver
browser.get('https://www.rpachallenge.com/')
browser.maximize_window()

#Fazendo download da planilha
p.sleep(1)
browser.find_element_by_xpath('/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button').click()
# p.alert('Pare')

#mudando formato da planilha para xlsx
p.sleep(1)
p.click(1355,139)
lendo_excel()




# #Fechando Navegador 
# p.sleep(2)
# browser.close()
# browser.quit()
