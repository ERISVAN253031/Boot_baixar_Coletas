from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service 
import pyautogui
import pandas as pd
import os
import time



navegador = webdriver.Chrome(ChromeDriverManager().install())

navegador.maximize_window()
navegador.get("https://airsupply.eslcloud.com.br/users/sign_in")

navegador.find_element('xpath','//*[@id="user_email"]').send_keys("email")
navegador.find_element('xpath','//*[@id="user_password"]').send_keys("senha")
navegador.find_element('xpath','//*[@id="new_user"]/div[2]/div/div[1]/input').click()

navegador.find_element('xpath','/html/body/div[2]/div/div/div[1]/ul[2]/li[3]/a').click()
time.sleep(7)
navegador.find_element('xpath','//*[@id="pick"]/a/span').click()
time.sleep(7)
navegador.find_element('xpath', "//*[text()='Convencional']").click()

df = pd.read_excel('COLETAS.xlsx')

camp_coleta = navegador.find_element('xpath','//*[@id="search_picks_sequence_code"]')


for linha in df['COLETAS']:
    linha = str(linha).strip()
    if len(linha) == 6:
        camp_coleta.clear()
        camp_coleta.send_keys(linha)
        time.sleep(2)
        navegador.find_element('xpath','//*[@id="search_picks_request_date"]').click()
        navegador.find_element('xpath','/html/body/div[11]/div[3]/div/button[2]').click()
        time.sleep(2)
        navegador.find_element('xpath','//*[@id="submit"]').click()
        time.sleep(2)
        navegador.find_element('xpath','//*[@id="picks_table"]/div[3]/div/table/tbody[1]/tr[1]/td[13]/div/div/button').click()
        navegador.find_element('xpath','//*[@id="picks_table"]/div[3]/div/table/tbody[1]/tr[1]/td[13]/div/div/ul/li[6]/a/span').click()
        time.sleep(3)
        pyautogui.click(870,760)
        time.sleep(6)
        navegador.find_element('xpath','//*[@id="picks_table"]/div[3]/div/table/tbody[1]/tr[1]/td[13]/div/div/button').click()
        navegador.find_element('xpath','//*[@id="picks_table"]/div[3]/div/table/tbody[1]/tr[1]/td[13]/div/div/ul/li[7]/a/span').click()
        time.sleep(3)
        pyautogui.click(870,760)

camp_coleta
camp_coleta.clear()        
camp_coleta.send_keys(linha)
time.sleep(2)

navegador.quit()
time.sleep(2)

pyautogui.click(40,1053)
time.sleep(1)
pyautogui.click(717,952)
time.sleep(1)
pyautogui.click(715,883)

#os.system("shutdown /s /t /0")
        
        

   