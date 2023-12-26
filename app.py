# script.py

from bs4 import BeautifulSoup
import openpyxl

def limpar_html_para_texto(html):
    # Analisa o HTML
    soup = BeautifulSoup(html, 'html.parser')
    
    # Obtém o texto limpo
    texto_limp = soup.get_text(separator=' ', strip=True)
    
    return texto_limp

# Carrega a planilha do Excel
caminho_planilha = 'bling-pyton.xlsm'
nome_da_planilha = 'Planilha'  # Substitua pelo nome da sua planilha
coluna_html = 'B'  # Substitua pela letra da coluna que contém o HTML

# Abre a planilha
wb = openpyxl.load_workbook(caminho_planilha)
planilha = wb[nome_da_planilha]

# Itera sobre as linhas e limpa o HTML na coluna específica
for linha in range(2, planilha.max_row + 1):
    celula_html = planilha[f'{coluna_html}{linha}']
    texto_limpo = limpar_html_para_texto(celula_html.value)
    print("Conteúdo antes:", celula_html.value)
    print("Conteúdo depois:", texto_limpo)
    celula_html.value = texto_limpo

# Salva as alterações na planilha
wb.save(caminho_planilha)
