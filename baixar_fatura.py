import os
import shutil
import regex as re
import pdfplumber
from openpyxl import load_workbook
import os
from dotenv import load_dotenv

load_dotenv()

def extrair_dados():
    # Pasta onde o arquivo foi criado
    pasta = os.getenv("PASTA_CRIACAO_ARQUIVO")
    
    # Lista os arquivos na pasta
    arquivos = os.listdir(pasta)

    # Encontre o arquivo mais recente na pasta
    arquivo_mais_recente = None
    data_modificacao = 0

    for arquivo in arquivos:
        caminho_arquivo = os.path.join(pasta, arquivo)
        if os.path.isfile(caminho_arquivo):
            data_arquivo = os.path.getmtime(caminho_arquivo)
            if data_arquivo > data_modificacao:
                data_modificacao = data_arquivo
                arquivo_mais_recente = caminho_arquivo

    if arquivo_mais_recente:
        # Agora você tem o caminho para o arquivo mais recente
        print(f"Arquivo mais recente: {arquivo_mais_recente}")        

        # Abra o arquivo PDF
        with pdfplumber.open(arquivo_mais_recente) as pdf:
            texto_completo = ''
            for page in pdf.pages:
                texto_completo += page.extract_text()

        print(texto_completo)
    else:
        print("Nenhum arquivo encontrado na pasta.")


    referencia = r'(?<=REFERÊNCIA:.*\n\d{2}\/\d{2}\/\d{4}\s\d+\s)(\d{2}\/\d{4})'
    referencia = re.findall(referencia, texto_completo)
    print(referencia[0])

    emissao = r'(?<=REFERÊNCIA:.*\n)(\d{2}\/\d{2}\/\d{4})'
    emissao = re.findall(emissao, texto_completo)
    print(emissao[0])

    vencimento = r'(?<=REFERÊNCIA:.*\n\d{2}\/\d{2}\/\d{4}\s\d+\s\d{2}\/\d{4}\s)(\d{2}\/\d{2}\/\d{4})'
    vencimento = re.findall(vencimento, texto_completo)
    print(vencimento)

    valor = r'(?<=REFERÊNCIA:.*\n.*R\$)(\d+\,\d+)'
    valor = re.findall(valor, texto_completo)
    print(valor)

    instalacao = r'(?<=GRANDE PAULISTA\/SP\s)(\d+)'
    instalacao = re.findall(instalacao, texto_completo)
    print(instalacao)

    path_dados = os.getenv("PATH_EXCEL_DADOS")
    workbook = load_workbook(path_dados)
    sheet = workbook['dados']
    valores = ['Enel', instalacao[0], referencia[0], emissao[0], vencimento[0], valor[0]]
    print(valores)
    sheet.append(valores)
    workbook.save(os.getenv("PATH_EXCEL_DADOS"))
    

if __name__ == "__main__":
    extrair_dados()