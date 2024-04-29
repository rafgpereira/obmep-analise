import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Função para contar a quantidade de medalhas por nível e tipo de medalha
def contar_medalhas_por_nivel(planilha):
    # Criando colunas para cada tipo de medalha
    planilha['Ouro'] = (planilha['Medalha'] == 'Ouro').astype(int)
    planilha['Prata'] = (planilha['Medalha'] == 'Prata').astype(int)
    planilha['Bronze'] = (planilha['Medalha'] == 'Bronze').astype(int)
    planilha['Mencao'] = (planilha['Menção'] == 'Sim').astype(int)
    
    # Agrupando os dados por nível e somando as medalhas por tipo
    contagem_por_nivel = planilha.groupby('Nível').agg({'Ouro': 'sum', 'Prata': 'sum', 'Bronze': 'sum', 'Mencao': 'sum'})
    return contagem_por_nivel

# Função para contar a quantidade de medalhas por nível e tipo de medalha em um grupo de escolas
def contar_medalhas_por_nivel_em_escolas(planilha, trechos_escolas):
    # Filtrando os dados para incluir apenas as escolas que contenham algum dos trechos desejados
    dados_filtrados = planilha[planilha['Escola'].str.contains('|'.join(trechos_escolas), case=False)].copy()
    
    # Criando colunas para cada tipo de medalha
    dados_filtrados['Ouro'] = (dados_filtrados['Medalha'] == 'Ouro').astype(int)
    dados_filtrados['Prata'] = (dados_filtrados['Medalha'] == 'Prata').astype(int)
    dados_filtrados['Bronze'] = (dados_filtrados['Medalha'] == 'Bronze').astype(int)
    dados_filtrados['Mencao'] = (dados_filtrados['Menção'] == 'Sim').astype(int)
    
    # Agrupando os dados por nível e somando as medalhas por tipo
    contagem_por_nivel = dados_filtrados.groupby('Nível').agg({'Ouro': 'sum', 'Prata': 'sum', 'Bronze': 'sum', 'Mencao': 'sum'})
    return contagem_por_nivel


def contar_medalhas_por_nivel_em_cef(planilha):
    # Filtrando os dados para incluir apenas as escolas que contenham algum dos trechos desejados
    dados_filtrados = planilha[planilha['Escola'].str.contains('CEF 213 DE SANTA MARIA')].copy()
    
    # Criando colunas para cada tipo de medalha
    dados_filtrados['Ouro'] = (dados_filtrados['Medalha'] == 'Ouro').astype(int)
    dados_filtrados['Prata'] = (dados_filtrados['Medalha'] == 'Prata').astype(int)
    dados_filtrados['Bronze'] = (dados_filtrados['Medalha'] == 'Bronze').astype(int)
    dados_filtrados['Mencao'] = (dados_filtrados['Menção'] == 'Sim').astype(int)
    
    # Agrupando os dados por nível e somando as medalhas por tipo
    contagem_por_nivel = dados_filtrados.groupby('Nível').agg({'Ouro': 'sum', 'Prata': 'sum', 'Bronze': 'sum', 'Mencao': 'sum'})
    return contagem_por_nivel

# Carregar os dados
caminho_planilha = 'obmep.xlsx'

# Trechos dos nomes das escolas desejadas
trechos_escolas = ['SANTA MARIA', 'SANTOS DUMONT', 'SARGENTO LIMA']

# Criar uma nova planilha
wb = Workbook()

# Selecionar a primeira aba (sheet)
sheet = wb.active
sheet.title = 'analise'

# Preencher a planilha com os dados de todos os anos
for ano in range(2005, 2024):
    if ano == 2020:
        continue

    dados = pd.read_excel(caminho_planilha, sheet_name=str(ano))



    # Somar a quantidade de medalhas por nível e tipo de medalha
    quantidade_medalhas_por_nivel = contar_medalhas_por_nivel(dados)

    # Preencher os dados na planilha
    sheet.append([f'Distrito Federal - {ano}'])
    for row in dataframe_to_rows(quantidade_medalhas_por_nivel, index=True, header=True):
        sheet.append(row)
    sheet.append([])  # Adicionar uma linha em branco entre os anos



    quantidade_medalhas_por_nivel_escolas = contar_medalhas_por_nivel_em_escolas(dados, trechos_escolas)

    # Preencher os dados na planilha
    sheet.append([f'Santa Maria - {ano}'])
    for row in dataframe_to_rows(quantidade_medalhas_por_nivel_escolas, index=True, header=True):
        sheet.append(row)
    sheet.append([])  # Adicionar uma linha em branco entre os anos



    quantidade_medalhas_por_nivel_em_cef = contar_medalhas_por_nivel_em_cef(dados)

    # Preencher os dados na planilha
    sheet.append([f'CEF 213 DE SANTA MARIA - {ano}'])
    for row in dataframe_to_rows(quantidade_medalhas_por_nivel_em_cef, index=True, header=True):
        sheet.append(row)
    sheet.append([])  # Adicionar uma linha em branco entre os anos
    sheet.append([])  # Adicionar uma linha em branco entre os anos
    sheet.append([])  # Adicionar uma linha em branco entre os anos

# Salvar a nova planilha
caminho_nova_planilha = 'analise-obmep.xlsx'
wb.save(caminho_nova_planilha)
