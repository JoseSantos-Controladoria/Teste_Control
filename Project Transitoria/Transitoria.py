import pandas as pd
                                                                                                                
Conta_custo = {
    "40372": {
        "nome": "PREMIOS E GRATIFICACOES",
        "natureza": [
            "C-PREMIOS/INCENTIVOS",
            "C-DO IT"
        ]
    },
    "40373": {
        "nome": "SEGURO DE VIDA EM GRUPO",
        "natureza": [
            "SEGURO DE VIDA"
        ]
    },
    "40375": {
        "nome": "ASSISTENCIA MEDICA / ODONTOLOGICA",
        "natureza": [
            "ASSISTENCIA MEDICA",
            "ASSISTENCIA MDICA",
            "ASSISTENCIA ODONTOLOGICA",
            "ASSISTENCIA ODONT"
        ]
    },
    "40391": {
        "nome": "MEDICINA OCUPACIONAL",
        "natureza": [
            "PCMSO"
        ]
    },
    "40541": {
        "nome": "PARCEIROS OPERACIONAIS",
        "natureza": [
            "C-COLIGADA / TAXA ADM"
        ]
    },
    "40402": {
        "nome": "LANCHES E REFEICOES",
        "natureza": [
            "C-ALIMENTACAO",
            "COFFEE BREAK"
        ]
    },
    "40403": {
        "nome": "VIAGENS E ESTADAS",
        "natureza": [
            "C-HOSPEDAGEM",
            "D-HOSPEDAGEM",
            "C-AEREO",
            "C-PASSAGEM RODOVIARIA"
        ]
    },
    "40408": {
        "nome": "FRETES E CARRETOS",
        "natureza": [
            "C-MOTOBOY",
            "C-TRANSPORTADORA",
            "COLIGADA / TRANSPORTADORA"
        ]
    },
    "40409": {
        "nome": "MATERIAL DE CONSUMO",
        "natureza": [
            "MATERIAL DE APOIO"
        ]
    },
    "40411": {
        "nome": "UNIFORMES E EQUIPAMENTOS DE SEGURANÇA",
        "natureza": [
            "UNIFORME/MATERIAL DE TRABALHO"
        ]
    },
    "40413": {
        "nome": "TELECOMUNICACOES E INTERNET",
        "natureza": [
            "C-TELEFONE CELULAR"
        ]
    },
    "40414": {
        "nome": "LOCACAO DE VEICULOS",
        "natureza": [
            "C-LOCACAO DE VEICULO",
            "D-LOCACAO DE VEICULO"
        ]
    },
    "40415": {
        "nome": "LOCACOES EM GERAL",
        "natureza": [
            "C-ALUGUEL DE EQUIPAMENTO",
            "LOCACAO DE SALA"
        ]
    },
    "40416": {
        "nome": "SERVIÇOS GRÁFICOS/EDIÇÃO/CRIAÇÃO",
        "natureza": [
            "EQUIPAMENTOS DE AUDIO",
            "C-MATERIAL GRAFICO"
         ]
    },
    "40419": {
        "nome": "MANUTENÇÃO E DESPESAS COM VEÍCULOS",
        "natureza": [
            "MANUTENCAO DE CARRO"
        ]
    },
    "40420": {
        "nome": "CONDUÇÃO, ESTACIONAMENTO E PEDÁGIOS",
        "natureza": [
            "C-TAXI",
            "C-PEDAGIO",
            "C-ESTACIONAMENTO"
        ]
    },
    "40421": {
        "nome": "CORREIOS E MALOTES",
        "natureza": [
            "C-CORREIO"
        ]
    },
    "40422": {
        "nome": "SERVIÇOS/MATERIAL DE INFORMÁTICA",
        "natureza": [
            "C-SERVIÇOS DE INFORMATICA",
            "D-SERVIÇOS DE INFORMATICA"
        ]
    },
    "40423": {
        "nome": "LEGAIS/CÓPIAS E AUTENTICAÇÕES",
        "natureza": [
            "C-AUTENTICAÇÕES E CÓPIAS"
        ]
    },
    "40426": {
        "nome": "CURSOS E TREINAMENTOS",
        "natureza": [
            "D-CURSO DE FORMAÇÃO PROFISSIONAL",
            "C-CURSO DE FORMAÇÃO PROFISSIONAL"
        ]
    },
    "40428": {
        "nome": "PROMOÇÕES FESTAS E EVENTOS",
        "natureza": [
            "C-EVENTOS E FEIRAS"
        ]
    },
    "40442": {
        "nome": "MANUTENÇÃO DAS INSTALAÇÕES/CELULARES",
        "natureza": [
            "C-Conservacao e Manutencao de equipamentos",
            "C-Conservção e Manutenção de equipamentos"
        ]
    },
    "40401": {
        "nome": "COMBUSTÍVEIS E LUBRIFICANTES",
        "natureza": [
            "C-COMBUSTÍVEL",
            "D-COMBUSTÍVEL"
        ]
    },
    "40445": {
        "nome": "RECRUTAMENTO E SELEÇÃO",
        "natureza": [
            "C-ANTECEDENTES",
            "D-ANTECEDENTES"
        ]
    },
     "50360": {
        "nome": "LEGAIS/CÓPIAS E AUTENTICAÇÕES",
        "natureza": [
            "C-AUTENTICAÇÕES E CÓPIAS"
        ]
    },
     "40448": {
        "nome": "APARELHOS CELULARES",
        "natureza": [
            "C-APARELHO CELULAR",
            "D-APARELHO CELULAR"
        ]
    }
}


conta_despesa = {
    "50159": {
        "despesa": "PREMIOS E GRATIFICACOES",
        "natureza": [
            "C-PREMIOS/INCENTIVOS",
            "D-PREMIOS/INCENTIVOS"
        ]
    },
    "50160": {
        "despesa": "SEGURO DE VIDA EM GRUPO",
        "natureza": [
            "SEGURO DE VIDA"
        ]
    },
    "50179": {
        "despesa": "ASSISTENCIA MEDICA E ODONTOLOGICA",
        "natureza": [
            "ASSISTNCIA MDICA",
            "ASSISTENCIA MEDICA",
            "ASSISTENCIA ODONTOLOGICA",
            "ASSISTENCIA ODONT"
        ]
    },
    "50201": {
        "despesa": "MANUTENÇÃO DE INSTALAÇÕES",
        "natureza": [
            "CONSERVACAO E MANUTE",
            "CONSERVACAO E MANUTENCAO",
            "D-CONSERVACAO E MANUTENCAO",
            "C-CONSERVACAO E MANUTENCAO"
        ]
    },
    "50205": {
        "despesa": "FRETES E CARRETOS",
        "natureza": [
            "D-MOTOBOY"
        ]
    },
    "50208": {
        "despesa": "CURSOS E TREINAMENTOS",
        "natureza": [
            "D-CURSOS DE FORMAÇÃO PROFISSIONAL"
        ]
    },
    "50210": {
        "despesa": "HONORÁRIOS ADVOCATÍCIOS",
        "natureza": [
            "D-HONORÁRIOS ADVOCATICIO"
        ]
    },
    "50211": {
        "despesa": "HONORÁRIOS CONTÁBEIS",
        "natureza": [
            "ASSESSORIA CONTÁBIL"
        ]
    },
    "50212": {
        "despesa": "RECRUTAMENTO E SELEÇÃO",
        "natureza": [
            "D-ANTECEDENTES"
        ]
    },
    "50213": {
        "despesa": "AUDITORIA E CONSULTORIA",
        "natureza": [
            "D-CONSULTORIA"
        ]
    },
    "50216": {
        "despesa": "SERVIÇOS GRÁFICOS/EDIÇÃO/CRIAÇÃO",
        "natureza": [
            "C-MATERIAL GRÁFICO",
            "C-EQUIPAMENTOS DE ÁUDIO"
        ]
    },
    "50217": {
        "despesa": "MEDICINA OCUPACIONAL",
        "natureza": [
            "PCMSO"
        ]
    },
    "50920": {
        "despesa": "SERVIÇOS PRESTADOS PJ'S",
        "natureza": [
            "D-MÃO OBRA TERCEIRO"
        ]
    },
    "50222": {
        "despesa": "LOCACAO DE EQUIPAMENTO PAGO A PJ",
        "natureza": [
            "D-ALUGUEL DE EQUIPAMENTOS"
        ]
    },
    "50223": {
        "despesa": "LOCACAO DE VEICULO",
        "natureza": [
            "D-LOCACAO DE VEICULO"
        ]
    },
    "50356": {
        "despesa": "TELECOMUNICAÇÕES E INTERNET",
        "natureza": [
            "C-TELEFONE CELULAR",
            "D-INTERNET",
            "C-INTERNET",
            "D-TELEFONE CELULAR",
            "TELEFONE FIXO"
        ]
    },
    "50360": {
        "despesa": "XEROX E AUTENTICAÇÕES",
        "natureza": [
            "D-AUTENTICAÇÕES E CÓPIAS"
        ]
    },
    "50364": {
        "despesa": "CORREIOS E MALOTES",
        "natureza": [
            "CORREIOS"
        ]
    },
    "50203": {
        "despesa": "MANUTENÇÃO DE VEÍCULOS",
        "natureza": [
            "MANUTENCAO DE CARRO"
        ]
    },
    "50244": {
        "despesa": "LICENCIAMENTO DE VEÍCULOS",
        "natureza": [
            "LICENCIAMENTO DE VEÍCULOS"
        ]
    },
    "50361": {
        "nome": "VIAGENS E ESTADAS",
        "natureza": [
            "C-HOSPEDAGEM",
            "D-HOSPEDAGEM",
            "C-AEREO",
            "C-PASSAGEM RODOVIARIA"
        ]
    },
    "50375": {
        "nome": "UNIFORMES E EQUIPAMENTOS DE SEGURANÇA",
        "natureza": [
            "UNIFORME/MATERIAL DE TRABALHO"
        ]
    },
    "50352": {
        "nome": "COMBUSTÍVEIS E LUBRIFICANTES",
        "natureza": [
            "C-COMBUSTÍVEL",
            "D-COMBUSTÍVEL"
        ]
    },
    "50388": {
        "nome": "PROMOÇÕES FESTAS E EVENTOS",
        "natureza": [
            "C-EVENTOS E FEIRAS"
        ]
     }
    
}

def normalize_string(s):
    """Normaliza a string para comparações consistentes."""
    return s.strip().upper()

def encontrar_conta(natureza, tipo):
    """Encontra a conta com base na natureza e tipo."""
    tipo = normalize_string(tipo)
    natureza = normalize_string(natureza)

    for conta, info in Conta_custo.items():
        if tipo == "CUSTOS" and natureza in info['natureza']:
            return conta

    for conta, info in conta_despesa.items():
        if tipo == "DESPESAS" and natureza in info['natureza']:
            return conta

    return None

def normalize(s):
    
    return s.strip().upper()

def encontrar_conta(natureza, tipo):
    
    tipo = normalize_string(tipo)
    natureza = normalize

def associar_conta(row):
    """Associa a conta ao DataFrame com base na natureza e tipo."""
    tipo = normalize_string(row['CUSTO OU DESPESA'])
    natureza = normalize_string(row['NATUREZA'])
    conta = encontrar_conta(natureza, tipo)
    
    if conta is None:
        print(f"Não foi possível encontrar conta para natureza: {natureza} e tipo: {tipo}")
    
    # Garantir que retornamos apenas um valor
    return conta if isinstance(conta, str) else None

def filtrar_e_associar_dados(caminho_entrada, valores_aps, salvar_em_arquivo, caminho_saida=None):
    """Filtra e associa dados com base nos valores de entrada."""
    df = pd.read_excel(caminho_entrada)
    
    print(f"Colunas disponíveis no arquivo: {df.columns.tolist()}")  # Debug: exibir colunas do DataFrame

    if 'APS' not in df.columns:
        raise ValueError("A coluna 'APS' não foi encontrada no arquivo.")

    df_filtrado = df[df['APS'].astype(str).isin(valores_aps)].copy()

    if 'NATUREZA' not in df_filtrado.columns:
        raise ValueError("A coluna 'NATUREZA' não foi encontrada no arquivo filtrado.")
    if 'CUSTO OU DESPESA' not in df_filtrado.columns:
        raise ValueError("A coluna 'CUSTO OU DESPESA' não foi encontrada no arquivo filtrado.")

    print(f"Antes da aplicação: {df_filtrado.head()}")  # Verifica o DataFrame filtrado antes da aplicação

    df_filtrado['CONTA DÉBITO'] = df_filtrado.apply(associar_conta, axis=1)

    print(f"Depois da aplicação: {df_filtrado.head()}")  # Verifica o DataFrame filtrado após a aplicação

    valores_nao_encontrados = [valor for valor in valores_aps if valor not in df['APS'].astype(str).values]
    valores_encontrados = [valor for valor in valores_aps if valor in df['APS'].astype(str).values]

    if valores_encontrados:
        print(f"Dados filtrados: {', '.join(valores_encontrados)}")
    if valores_nao_encontrados:
        print(f"Valores não encontrados: {', '.join(valores_nao_encontrados)}")

    print("Dados filtrados e associados:")
    print(df_filtrado.head())

    if salvar_em_arquivo:
        if not caminho_saida:
            caminho_saida = 'planilha_associada.xlsx'
        df_filtrado.to_excel(caminho_saida, index=False, engine='openpyxl')
        print(f"Dados filtrados e associados salvos em: {caminho_saida}")
    else:
        print("Dados filtrados e associados não foram salvos em um arquivo.")

    return df_filtrado

# Solicitando inputs do usuário
caminho_entrada = input("Digite o caminho para o arquivo Excel de entrada: ")

# Entrada do usuário para múltiplos valores na coluna 'APS'
valores_aps_str = input("Digite os valores que deseja filtrar na coluna 'APS', separados por vírgula: ")
valores_aps = [valor.strip() for valor in valores_aps_str.split(',')]

salvar_em_arquivo = input("Deseja salvar os resultados em um arquivo Excel? (s/n): ").strip().lower()

caminho_saida = None
if salvar_em_arquivo == 's':
    caminho_saida = input("Digite o caminho para o arquivo Excel de saída (deixe em branco para usar 'planilha_associada.xlsx'): ")
    caminho_saida = caminho_saida if caminho_saida else 'planilha_associada.xlsx'

# Filtrar e associar as contas
resultado = filtrar_e_associar_dados(caminho_entrada, valores_aps, salvar_em_arquivo == 's', caminho_saida)
print("Dados filtrados e associados:")
print(resultado)

