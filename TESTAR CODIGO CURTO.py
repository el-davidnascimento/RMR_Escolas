from datetime import datetime
import openpyxl
import pandas as pd
import datetime as date

############################################# 12 - Aderencia_Totvs ##########################################################
############################################# Tabela Aderencia Totvs
############################# INDICADOR : Aderência do Totvs

caminho_arquivo_totvs = r'G:\.shortcut-targets-by-id\1kArAZwgCxrjbQwQOPEzeJLtMUll3VVJ7\13. Dados\13.2. RMR\13.2.2. Escolas\13.2.2.1. UNI\13.2.2.1.2. Quantitativo\13.2.2.1.2.1. Base\13.2.2.1.2.1.2. Base Escolas\12 - Aderencia_Totvs.xlsx'
nome_aba_totvs = '16 - Aderência Totvs'
dados_excel_totvs = pd.read_excel(caminho_arquivo_totvs, sheet_name=nome_aba_totvs)

# Criar a lista de meses de 1 a 12
meses = list(range(1, 13))
ano_atual = datetime.now().year

# Carregar o arquivo de métrica
planilha = openpyxl.load_workbook(r'G:\.shortcut-targets-by-id\1kArAZwgCxrjbQwQOPEzeJLtMUll3VVJ7\13. Dados\13.2. RMR\13.2.2. Escolas\13.2.2.1. UNI\13.2.2.1.3. Métricas\1. Novos Indicadores Escola.xlsx')

# Selecionar a planilha desejada
Aba_metricas = planilha['Planilha1']

# Puxar o valor da célula A1
Valor_frente1_aderencia_totvs = Aba_metricas['A16'].value
Valor_area_responsavel_aderencia_totvs = Aba_metricas['B16'].value
Valor_setor_responsavel_aderencia_totvs = Aba_metricas['C16'].value
Valor_frente2_aderencia_totvs = Aba_metricas['D16'].value
Valor_indicador_aderencia_totvs = Aba_metricas['E16'].value
Valor_meta_aderencia_totvs = Aba_metricas['F16'].value
Valor_ruim_aderencia_totvs = Aba_metricas['G16'].value
Valor_regular_aderencia_totvs = Aba_metricas['H16'].value
Valor_otimo_aderencia_totvs = Aba_metricas['I16'].value
Valor_peso_aderencia_totvs = Aba_metricas['J16'].value

# Criar os dados para a tabela
dados = {
    'Ano': [ano_atual] * len(meses),
    'Mês': meses,
    'Unidade': ["UNI"] * len(meses),
    'FRENTE 1': Valor_frente1_aderencia_totvs,
    'ÁREA RESPONSÁVEL RESULTADO': Valor_area_responsavel_aderencia_totvs,
    'SETOR RESPONSÁVEL COLETA DADO': Valor_setor_responsavel_aderencia_totvs,
    'FRENTE 2': Valor_frente2_aderencia_totvs,
    'INDICADOR': Valor_indicador_aderencia_totvs,
    'META': Valor_meta_aderencia_totvs,
    'RUIM': Valor_ruim_aderencia_totvs,
    'REGULAR': Valor_regular_aderencia_totvs,
    'ÓTIMO': Valor_otimo_aderencia_totvs,
    'Valor': [0] * len(meses),
    'Tipo':["Porcentagem"] * len(meses),
    'Peso': Valor_peso_aderencia_totvs,
}

# Criar o DataFrame
df_totvs = pd.DataFrame(dados)
qtde_implantada = 0

# Inicializar as contagens
qtde_implantada = 0
qtde_total_etapas = 0

# Iterar sobre as linhas do DataFrame 'dados_excel_totvs'
for index, row in dados_excel_totvs.iterrows():
    if row['STATUS'] == 'Implantado':
        qtde_implantada += 1
        qtde_total_etapas += 1
    else:
        qtde_total_etapas += 1

# Calcular a proporção de etapas concluídas
proporcao_concluida = qtde_implantada / qtde_total_etapas

df_totvs['Valor'] = proporcao_concluida

df_totvs['Valor'] = df_totvs['Valor'].map('{:.2%}'.format)

Tratamento_totvs = r'G:\.shortcut-targets-by-id\1kArAZwgCxrjbQwQOPEzeJLtMUll3VVJ7\13. Dados\13.2. RMR\13.2.2. Escolas\13.2.2.1. UNI\13.2.2.1.2. Quantitativo\13.2.2.1.2.1. Base\13.2.2.1.2.1.2. Base Escolas\13.2.2.1.2.1.2.3. Excel\Tratamento_totvs.xlsx'
df_totvs.to_excel(Tratamento_totvs,index=False)
