import openpyxl

# Defina o caminho do arquivo Excel ou substitua pelo caminho do seu arquivo
caminho_arquivo = 'completo_01-08-2023.xlsx'

# Carregue o arquivo Excel usando a função load_workbook do openpyxl
arquivo_excel = openpyxl.load_workbook(caminho_arquivo)

# Escolha a planilha que deseja ler (caso haja várias planilhas no arquivo)
nome_planilha = 'Report'
planilha = arquivo_excel[nome_planilha]

# Verifique se a coluna 'Códigos' existe na planilha
coluna_codigos = planilha['A']  # Supondo que a coluna 'Códigos' esteja na coluna C

if coluna_codigos:
    # Acesse os valores da coluna 'Códigos' e imprima-os
    codigos = [celula.value for celula in coluna_codigos if celula.value is not None]

    # Imprima os itens da coluna 'Códigos'
    print(codigos)
else:
    print("A coluna 'Códigos' não existe na planilha.")

