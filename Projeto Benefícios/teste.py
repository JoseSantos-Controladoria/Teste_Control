import pandas as pd

# Caminho correto para o arquivo
input_file = r'C:/Users/jose.santos/Desktop/Automacoes/Projeto Benefícios/base_beneficio.xlsx'  # Alterar para o nome real do arquivo

# Carregar a planilha base
df_base = pd.read_excel(input_file)

# Remover espaços extras nos nomes das colunas
df_base.columns = df_base.columns.str.strip()

# Imprimir as colunas para garantir que estão corretas
print("Colunas da planilha:", df_base.columns)

# Lista das colunas de benefícios (todas as colunas, exceto 'CC')
beneficios = ['VR', 'CB', 'VA', 'OB', 'AJ', 'VT', 'PJ', 'AUX MORADIA', 'DESPESAS ACIONISTAS', 'Taxa']

# Criar um objeto ExcelWriter para salvar o novo arquivo
output_file = r'C:/Users/jose.santos/Desktop/Automacoes/Projeto Benefícios/beneficios_filtrados_agregados.xlsx'  # Nome do arquivo de saída

with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    # Iterar por cada benefício
    for beneficio in beneficios:
        # Verificar se a coluna existe na planilha
        if beneficio in df_base.columns:
            # Fazer uma cópia da planilha base para não alterar os dados
            df_temp = df_base.copy()

            # Substituir o valor "          -  " (com espaços) por NaN para garantir que será removido
            df_temp[beneficio] = df_temp[beneficio].replace("          -  ", pd.NA)

            # Converter os valores da coluna de benefício para numérico, erros serão convertidos para NaN
            df_temp[beneficio] = pd.to_numeric(df_temp[beneficio], errors='coerce')

            # Filtrar os dados da coluna específica para remover valores vazios, nulos e zero
            df_temp = df_temp[df_temp[beneficio].notna() & (df_temp[beneficio] != 0)]

            # Agrupar os dados por 'CC' e 'NOME ARQUIVO' e somar os valores
            df_aggregated = df_temp.groupby(['CC', 'NOME ARQUIVO'])[beneficio].sum().reset_index()

            # Salvar os dados agregados em uma nova aba (sheet)
            df_aggregated.to_excel(writer, sheet_name=beneficio, index=False)
        else:
            print(f"A coluna '{beneficio}' não foi encontrada na planilha.")

print(f"Planilha criada com sucesso em: {output_file}")
