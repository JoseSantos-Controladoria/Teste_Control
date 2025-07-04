import pandas as pd
from colorama import Fore, Style, init

init(autoreset=True)

CAMINHO_ARQUIVO = 'C:/Users/jose.santos/Desktop/Automacoes/Projeto Filtragem de Dados/Filtro/base_de_aps.xlsx'

def filtrar_dados(caminho_entrada, valores_aps, salvar_em_arquivo, caminho_saida=None):

    df = pd.read_excel(caminho_entrada)

    print(Fore.YELLOW + "Cabeçalhos das Colunas:")
    print(Fore.GREEN + str(df.columns))
    print(Fore.YELLOW + "Primeiras Linhas do DataFrame:")
    print(Fore.GREEN + str(df.head()))

    if 'APS' not in df.columns:
        raise ValueError(Fore.RED + "A coluna 'APS' não foi encontrada no arquivo.")

    df['APS'] = df['APS'].astype(str)
    valores_aps = [valor.strip() for valor in valores_aps]

    df_filtrado = df[df['APS'].isin(valores_aps)]

    valores_nao_encontrados = [valor for valor in valores_aps if valor not in df['APS'].values]
    valores_encontrados = [valor for valor in valores_aps if valor in df['APS'].values]

    indice_coluna_aps = df.columns.get_loc('APS')
    colunas_resultado = df.columns[indice_coluna_aps:]
    df_resultado = df_filtrado[colunas_resultado]

    if valores_encontrados:
        print(Fore.GREEN + f"Dados filtrados: {', '.join(valores_encontrados)}")
    if valores_nao_encontrados:
        print(Fore.RED + f"Valores não encontrados: {', '.join(valores_nao_encontrados)}")

    if salvar_em_arquivo:
        if not caminho_saida:
            caminho_saida = 'planilha_filtrada.xlsx'
        df_resultado.to_excel(caminho_saida, index=False)
        print(Fore.CYAN + f"Dados filtrados salvos em: {caminho_saida}")
    else:
        print(Fore.YELLOW + "Dados filtrados não foram salvos em um arquivo.")

    return df_resultado



def consultar_aps(caminho_entrada, aps_consulta):
    df = pd.read_excel(caminho_entrada)

    if 'APS' not in df.columns:
        raise ValueError(Fore.RED + "A coluna 'APS' não foi encontrada no arquivo.")

    df['APS'] = df['APS'].astype(str)

    df_consulta = df[df['APS'].str.contains(aps_consulta, case=False, na=False)]

    if not df_consulta.empty:
        print(Fore.GREEN + f"Resultados da consulta para 'APS' '{aps_consulta}':")
        print(Fore.GREEN + str(df_consulta))
    else:
        print(Fore.RED + f"Nenhum resultado encontrado para 'APS' '{aps_consulta}'.")

def menu():
    while True:
        print(Fore.CYAN + "\nMenu:")
        print(Fore.CYAN + "1 - Realizar o filtro")
        print(Fore.CYAN + "2 - Consultar uma APS")
        print(Fore.CYAN + "3 - Sair")
        
        escolha = input(Fore.BLUE + "Escolha uma opção (1/2/3): " + Style.RESET_ALL).strip()

        if escolha == '1':
            valores_aps_str = input(Fore.BLUE + "Digite os valores que deseja filtrar na coluna 'APS', separados por vírgula: " + Style.RESET_ALL)
            valores_aps = [valor.strip() for valor in valores_aps_str.split(',')]
            salvar_em_arquivo = input(Fore.BLUE + "Deseja salvar os resultados em um arquivo Excel? (s/n): " + Style.RESET_ALL).strip().lower()
            caminho_saida = None
            if salvar_em_arquivo == 's':
                caminho_saida = input(Fore.BLUE + "Digite o caminho para o arquivo Excel de saída (deixe em branco para usar 'planilha_filtrada.xlsx'): " + Style.RESET_ALL)
                caminho_saida = caminho_saida if caminho_saida else 'planilha_filtrada.xlsx'

            resultado = filtrar_dados(CAMINHO_ARQUIVO, valores_aps, salvar_em_arquivo == 's', caminho_saida)
            print(Fore.YELLOW + "Dados filtrados:")
            print(Fore.GREEN + str(resultado))

        elif escolha == '2':
            aps_consulta = input(Fore.BLUE + "Digite o valor da APS que deseja consultar: " + Style.RESET_ALL)
            consultar_aps(CAMINHO_ARQUIVO, aps_consulta)
        
        elif escolha == '3':
            print(Fore.GREEN + "Saindo do programa. Até logo!")
            break
        
        else:
            print(Fore.RED + "Opção inválida. Tente novamente.")

if __name__ == "__main__":
    menu()
