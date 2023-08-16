import openpyxl

# Carregar o arquivo Excel
workbook = openpyxl.load_workbook('tabelas\completo_01-08-2023.xlsx')
sheet = workbook.active

# Dicionário para armazenar os dados filtrados por link2
filtered_data = {}

# Iterar pelas linhas da planilha (começando da segunda linha, assumindo que a primeira linha é o cabeçalho)
for row in sheet.iter_rows(min_row=2, values_only=True):
    codigo = row[0]
    criado_em = row[8]
    link2 = row[23]
    if link2 not in filtered_data:
        filtered_data[link2] = []
    
    # Verificar se o chamado é do mês 07/2023
    if criado_em.month == 7 and criado_em.year == 2023:
        chamado = {
            'Código': codigo,
            'Criado em': criado_em,
            'Título do problema': row[16],
            'Descrição do problema': row[17],
            'Data e hora que retornou ao normal': row[30],
            'Observações finais': row[31],
            'Laudo Final': row[54],
            'Tempo de atraso do cliente em horas': row[58],
            'Tempo de atraso do cliente em minutos': row[59],
            'Justificativa do atraso do cliente': row[61]
        }
        filtered_data[link2].append(chamado)

# Imprimir os resultados formatados
for link2, chamados in filtered_data.items():
    print(f"{link2}")
    for chamado in chamados:
        print(f"\tCódigo {chamado['Código']} (relacionado ao {link2}):")
        print(f"\t\tCriado em: {chamado['Criado em']}")
        print(f"\t\tTítulo do problema: {chamado['Título do problema']}")
        print(f"\t\tDescrição do problema: {chamado['Descrição do problema']}")
        print(f"\t\tData e hora que retornou ao normal: {chamado['Data e hora que retornou ao normal']}")
        print(f"\t\tObservações finais: {chamado['Observações finais']}")
        print(f"\t\tLaudo Final: {chamado['Laudo Final']}")
        print(f"\t\tTempo de atraso do cliente em horas: {chamado['Tempo de atraso do cliente em horas']}")
        print(f"\t\tTempo de atraso do cliente em minutos: {chamado['Tempo de atraso do cliente em minutos']}")
        print(f"\t\tJustificativa do atraso do cliente: {chamado['Justificativa do atraso do cliente']}")
        print('----------------------------------------------------------------------------------------------------------------------------------')
        print('==================================================================================================================================')