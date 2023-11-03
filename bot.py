import time
import pandas as pd
import openpyxl
import datetime
import pygetwindow as gw
from send_email import send_email

data_atual = datetime.date.today()
mes_ano = data_atual.strftime("%m-%Y")
hora_minuto = datetime.datetime.now().strftime('%H:%M')

if __name__ == "__main__":
    subject = 'RPA - Conta de Consumo'    
    body = 'Processo iniciado'
    attachment_path = ''
    send_email(subject, body, attachment_path)


planilha_original = pd.read_excel('Relatório de execução.xlsx')
relatorio_atual = planilha_original.copy()
nome_relatorio = "Relatório de execução " + mes_ano +".xlsx"
relatorio_atual.to_excel(nome_relatorio, index=False)

workbook = openpyxl.load_workbook('login.xlsx')
sheet_produtos = workbook['login']
quantidade_de_linhas = sheet_produtos.max_row

#Define resolução da tela
'''
largura, altura = 1920, 1080
janela_principal = gw.getWindowsWithTitle('')[0]
janela_principal.size = (largura, altura)
'''

for linha in sheet_produtos.iter_rows(min_row=2, values_only=True):
    print(linha[0])
    #Continuar daqui para baixo, implementar logica as filiais
    hora_minuto = datetime.datetime.now().strftime('%H:%M')
    nome_relatorio = "Relatório de execução " + mes_ano +".xlsx"

if __name__ == "__main__":
    subject = 'RPA - Conta de Consumo'    
    body = 'Processo Finalizado'
    attachment_path = "Relatório de execução " + mes_ano +".xlsx"
    send_email(subject, body,attachment_path)