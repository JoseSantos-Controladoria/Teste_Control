import pandas as pd
import pyodbc
import warnings
import time
import re
import xlsxwriter

warnings.filterwarnings('ignore')

pd.set_option('display.max_columns', None)

#server = "186.231.24.138,10252"
server = "192.168.26.22,10252"
db = "INTEGRA_WORKON"
username = "giselecao"
password = "Giselecao@123"  
conn = pyodbc.connect('driver={SQL Server};SERVER='+ server + ';DATABASE=' + db + ';UID='+username+';PWD='+ password)


print('Insira a competência inicio e fim para baixar as aps:')
ano = input('Insira o Ano: ')
mes = input('Insira o Mês: ')
if len(mes) == 1:
    mes = '0'+mes
cpt_ini = ano + mes


ano = input('Insira o Ano: ')
mes = input('Insira o Mês: ')
if len(mes) == 1:
    mes = '0'+mes
cpt_fim = ano + mes


sql_sup = """SELECT
 TB_Titulo.NumeroTituloPrincipal
,ClienteFornec                          AS sup_ClienteFornec
,TB_Empresa.cgc 						AS sup_cnpj
,TB_Titulo.CodigoEmpresa				AS sup_CodigoEmpresa
,CodigoBanco							AS sup_CodigoBanco
,Status									AS sup_Status
,DataBaixa								AS sup_DataBaixa
,DataCompetencia						AS sup_DataCompetencia
,ValorTitulo							AS sup_ValorTitulo
,ValorJuros								AS sup_ValorJuros
,ValorDesconto							AS sup_ValorDesconto
,ValorMulta								AS sup_ValorMulta
,Duplicata								AS sup_Duplicata
,ValorDesctoAdtoForn					AS sup_ValorDesctoAdtoForn
,TB_Titulo.NumeroTituloPrincipal		AS sup_NumeroTituloPrincipal
,NumeroAP								AS sup_NumeroAp
,TB_Titulo.Cheque						AS sup_Cheque
,Rateio									AS sup_valid_rateio
,DescHistorico							AS SUP_DescHistorico
,CodigoCentroCusto						AS SUP_CodigoCentroCusto
,Portador								AS SUP_Portador
,CodigoDespesa							AS SUP_CodigoDespesa
,Descricao								AS SUP_Descricao



FROM TB_Titulo 
LEFT JOIN TB_TipoDespesa ON TB_TipoDespesa.CodigoDespesa = TB_Titulo.Despesa
LEFT JOIN TB_Empresa on CONCAT(tb_empresa.CodigoEmpresa, tb_empresa.CodigoFilial) = CONCAT(TB_Titulo.CodigoEmpresa, TB_Titulo.CodigoFilial)
WHERE 
STATUS = 'B' AND
CAST(DataCompetencia as INT) between {} and {} AND 
NumeroTituloPrincipal = '' AND
NumeroAP != 0 """.format(cpt_ini,cpt_fim)

df_sup = pd.io.sql.read_sql(sql_sup,conn)

aps = df_sup['sup_NumeroAp'].to_list()
aps = str(aps)
aps = aps.replace('[','(')
aps = aps.replace(']',')')

sql_inf = """
SELECT
 DescHistorico										AS INF_DescHistorico
,CodigoCentroCusto									AS INF_CodigoCentroCusto
,DataCompetencia									AS INF_DataCompetencia
,ValorTitulo										AS INF_ValorTitulo
,Portador											AS INF_Portador
,NumeroAP											AS INF_NumeroAP
,NumeroAPPrestacaoContas							AS INF_NumeroAPPrestacaoContas
,CodigoDespesa										AS INF_CodigoDespesa
,Descricao											AS INF_Descricao
,TB_Titulo.Cheque									AS INF_Cheque
,TB_Titulo.NumeroTituloPrincipal					AS inf_NumeroTituloPrincipal
,CASE 
		WHEN NumeroTituloPrincipal = '' then  'LINHA PRINCIPAL - SUPERIOR'
		ELSE 'LINHA RATEIO'
END 												AS TIPO_DADO

FROM TB_Titulo 
LEFT JOIN TB_TipoDespesa ON TB_TipoDespesa.CodigoDespesa = TB_Titulo.Despesa
where
STATUS = 'R' AND
NumeroAP IN {}
AND CodigoBancoOri != 999
AND NumeroTituloPrincipal != ''

""".format(aps)

df_inf = pd.io.sql.read_sql(sql_inf,conn)

df = pd.merge(df_sup, df_inf, how = 'left', left_on = 'sup_NumeroAp', right_on = 'INF_NumeroAP')

df_ok = df
df_ok = df_ok[df_ok['INF_NumeroAP'].apply(str) != 'nan']
df_nok = df[df['INF_NumeroAP'].apply(str) == 'nan']

df_nok['INF_DescHistorico'] = df_nok['SUP_DescHistorico']
df_nok['INF_CodigoCentroCusto'] = df_nok['SUP_CodigoCentroCusto']
df_nok['INF_DataCompetencia'] = df_nok['sup_DataCompetencia']
df_nok['INF_ValorTitulo'] = df_nok['sup_ValorTitulo']
df_nok['INF_Portador'] = df_nok['SUP_Portador']
df_nok['INF_NumeroAP'] = df_nok['sup_NumeroAp']
df_nok['INF_Cheque'] = df_nok['sup_Cheque']
df_nok['inf_NumeroTituloPrincipal'] = df_nok['sup_NumeroTituloPrincipal']


df_final = pd.concat([df_nok,df_ok])

# Função para remover caracteres problemáticos, incluindo o "▼"
def clean_text_directly(text):
    if isinstance(text, str):
        # Remoção direta do caractere problemático e outros caracteres não ASCII
        cleaned_text = text.replace('▼', '')
        cleaned_text = re.sub(r'[^\x00-\x7F]+', '', cleaned_text)
        return cleaned_text
    else:
        return text

# Verificar se alguma célula contém o caractere antes da limpeza
def check_for_character(df, character='▼'):
    for col in df.columns:
        if df[col].dtype == 'object':  # Se a coluna for de texto
            contains_char = df[col].apply(lambda x: character in x if isinstance(x, str) else False)
            if contains_char.any():
                print(f"Caractere '{character}' encontrado na coluna: {col}")
                return True
    return False

# Aplica a limpeza diretamente
for col in df_final.columns:
    df_final[col] = df_final[col].apply(clean_text_directly)

# Verifica se ainda existe o caractere no DataFrame
if check_for_character(df_final):
    print("Ainda existem células com o caractere problemático.")
else:
    print("Nenhuma célula com o caractere problemático encontrada.")

# Tentativa de salvar imediatamente após a limpeza
try:
    df_final.to_excel('aps.xlsx', index=False)
    print("Arquivo salvo com sucesso!")
except Exception as e:
    print(f"Erro ao salvar o arquivo: {e}")

# Limpa os nomes das colunas
df_final.columns = [clean_text_directly(col) for col in df_final.columns]

# Tentativa de salvar após limpar os nomes das colunas
try:
    df_final.to_excel('aps_corrigido.xlsx', index=False)
    print("Arquivo salvo com sucesso!")
except Exception as e:
    print(f"Erro ao salvar o arquivo: {e}")
try:
    df_final.to_csv('aps.txt', sep='\t', index=False)
    print("Arquivo salvo com sucesso!")
except Exception as e:
    print(f"Erro ao salvar o arquivo: {e}")

# Cria um novo arquivo Excel e adiciona uma planilha
workbook = xlsxwriter.Workbook('aps_alt2.xlsx', {'nan_inf_to_errors': True})
worksheet = workbook.add_worksheet()

# Preenche a planilha com os dados do DataFrame
for row_num, (index, row) in enumerate(df_final.iterrows()):
    for col_num, value in enumerate(row):
        worksheet.write(row_num, col_num, value)

workbook.close()
