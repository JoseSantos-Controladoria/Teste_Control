import pandas as pd
import os
import zipfile

caminho = r"C:\Users\jose.santos\Desktop\Automacoes\Benefícios.py"

nome_pasta = os.path.basename(caminho)

novo_diretorio = os.path.join(caminho, "CONCAT")
if not os.path.exists(novo_diretorio):
    os.makedirs(novo_diretorio)

arquivos = os.listdir(caminho)

arquivos_xlsx = [arquivo for arquivo in arquivos if arquivo.endswith(".xlsx")]

dfs = []
for arquivo in arquivos_xlsx:
    caminho_arquivo = os.path.join(caminho, arquivo)
    try:
        df = pd.read_excel(caminho_arquivo, engine="openpyxl", index_col=None, header=0)
        dfs.append(df)
    except zipfile.BadZipFile:
        print(f"Error reading file: {caminho_arquivo}. The file may be corrupted or not a valid Excel file.")

if len(dfs) == 0:
    print("No valid Excel files found.")

else:
    # Adiciona uma nova coluna ao dataframe com o nome do arquivo original
    for i, df in enumerate(dfs):
        arquivo = arquivos_xlsx[i]
        df['Arquivo Original'] = arquivo

    # Concatena os dataframes da lista "dfs" em um dataframe final
    df_final = pd.concat(dfs, ignore_index=True)

    nome_arquivo_saida = input("Digite o nome do arquivo de saída (sem a extensão): ")

    caminho_saida = os.path.join(novo_diretorio, f"{nome_arquivo_saida}.xlsx")
    df_final.to_excel(caminho_saida, index=False)

    print(f"O arquivo foi salvo em {os.path.abspath(novo_diretorio)}.")
