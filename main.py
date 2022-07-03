from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
import time


def getObject():
    options = Options()
    options.add_argument("--disable-popup-blocking")
    options.add_argument("headless")
    browser = webdriver.Chrome(options=options)
    browser.get('https://bnccompras.com/Process/ProcessSearchActivity')
    time.sleep(2)

    atividade = browser.find_element(By.XPATH, '//*[@id="fkActivity"]')
    select_atividade = Select(atividade)
    select_atividade.select_by_value('128')

    for x in range(1, 28):
        try:

            estado = browser.find_element(By.CSS_SELECTOR, '#fkState')
            select_estado = Select(estado)
            select_estado.select_by_value('{}'.format(x))

            browser.find_element(By.XPATH, '//*[@id="btnAuctionSearch"]').click()

            time.sleep(2)

            objeto = browser.find_element(By.XPATH, '//*[@id="tableProcessDataBody"]/tr/td[4]').text
            print(objeto)

            browser.find_element(By.XPATH, '//*[@id="tableProcessDataBody"]/tr/td[1]/a').click()

        except:
            print('Nada novo')


getObject()
