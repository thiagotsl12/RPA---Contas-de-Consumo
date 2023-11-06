import time
import pandas as pd
import openpyxl
from dotenv import load_dotenv
import datetime
import pygetwindow as gw
import pyautogui
import os
from send_email import send_email
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

load_dotenv()

data_atual = datetime.date.today()
mes_ano = data_atual.strftime("%m-%Y")
hora_minuto = datetime.datetime.now().strftime('%H:%M')

'''if __name__ == "__main__":
    subject = 'RPA - Conta de Consumo'    
    body = 'Processo iniciado'
    attachment_path = ''
    send_email(subject, body, attachment_path)
'''

planilha_original = pd.read_excel('C:\\Users\\thiag\\OneDrive\\Documentos\\Hase tech\\Cases\\github\\RPA - Contas de Consumo\\RPA---Contas-de-Consumo\\Relatório de execução.xlsx')
relatorio_atual = planilha_original.copy()
nome_relatorio = "Relatório de execução " + mes_ano +".xlsx"
relatorio_atual.to_excel(nome_relatorio, index=False)

workbook = openpyxl.load_workbook('C:\\Users\\thiag\\OneDrive\\Documentos\\Hase tech\\Cases\\github\\RPA - Contas de Consumo\\RPA---Contas-de-Consumo\\login.xlsx')
sheet_produtos = workbook['login']
quantidade_de_linhas = sheet_produtos.max_row

#Define resolução da tela

largura, altura = 1920, 1080
janela_principal = gw.getWindowsWithTitle('')[0]
janela_principal.size = (largura, altura)

 # Configurações do Chrome
chrome_options = ChromeOptions()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--disable-popup-blocking")
chrome_options.add_argument("--disable-notifications")
chrome_options.add_argument("--disable-infobars")
# Abrir o Chrome
url = os.getenv("URL")
browser = webdriver.Chrome(options=chrome_options)
url = url  # Substitua pelo URL desejado
browser.get(url)
wait = WebDriverWait(browser, 30)

time.sleep(20)
pyautogui.click(960,596,duration=0.2)


# Clique no elemento

for linha in sheet_produtos.iter_rows(min_row=2, values_only=True):

    element = wait.until(EC.visibility_of_element_located((By.ID, 'email'))).send_keys(linha[0])
    element = wait.until(EC.visibility_of_element_located((By.ID, 'senha'))).send_keys(linha[1])
    elemento_submit = wait.until(EC.presence_of_element_located((By.XPATH, f'//*[@translate="@APP-LOGIN-ENTRAR"]'))).click()
    elemento_submit = wait.until(EC.presence_of_element_located((By.XPATH, f'//*[translate="@APP-COMMON-ENTRAR""]'))).click()
    #Continuar daqui para baixo, implementar logica as filiais
    hora_minuto = datetime.datetime.now().strftime('%H:%M')
    nome_relatorio = "Relatório de execução " + mes_ano +".xlsx"

'''if __name__ == "__main__":
    subject = 'RPA - Conta de Consumo'    
    body = 'Processo Finalizado'
    attachment_path = "Relatório de execução " + mes_ano +".xlsx"
    send_email(subject, body,attachment_path)
'''