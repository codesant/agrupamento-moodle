#

# Importação das bibliotecas
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
import datetime
import time
import openpyxl as excel
import urllib.parse
import os
import timeit
import random

#Inicia e salva os cookies
options = webdriver.ChromeOptions()
options.add_argument('--user-data-dir=./User_Data_MDI')
driver = webdriver.Chrome(chrome_options=options)
driver.get('https://ead.senaibahia.com.br/group/index.php?id=280')

input()
time.sleep(2)
# Itera sobre as linhas da planilha, lendo um contato de cada vez
file = excel.load_workbook('./contacts-message.xlsx')
sheet = file.active

for row in sheet.iter_rows(min_row=2, max_col=1):
    contact = str(row[0].value)

    # Procura o elemento que contém o texto do CPF
    cpf_element = None
    try:
        cpf_element = driver.find_element(By.XPATH, f"//*[contains(text(), '{contact}')]")
    except NoSuchElementException:
        pass

    # Clica no elemento se ele for encontrado
    if cpf_element:
        actions = ActionChains(driver)
        actions.key_down(Keys.COMMAND)
        actions.click(cpf_element)
        actions.key_up(Keys.COMMAND)
        actions.perform()

    print("Incluindo ", contact)

# Clica no botão de adicionar contatos
driver.find_element(By.XPATH, '//*[@id="add"]').click()
print("Concluído")


