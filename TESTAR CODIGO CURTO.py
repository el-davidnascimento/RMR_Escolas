from datetime import datetime
import openpyxl
import pandas as pd
import datetime as date

############################################# 3 - Desistencia_de_alunos ##########################################################

caminho_arquivo = r'G:\.shortcut-targets-by-id\1kArAZwgCxrjbQwQOPEzeJLtMUll3VVJ7\13. Dados\13.2. RMR\13.2.2. Escolas\13.2.2.1. UNI\13.2.2.1.2. Quantitativo\13.2.2.1.2.1. Base\13.2.2.1.2.1.2. Base Escolas\3 - Desistencia_de_alunos.xlsx'
nome_aba = '3 - Desistencia de alunos'
dados_excel = pd.read_excel(caminho_arquivo, sheet_name=nome_aba)

# Obter a data atual
data_atual = date.datetime.today()

# Criar uma nova coluna com a data atual
dados_excel['Date_insert'] = data_atual
dados_excel.loc[dados_excel['Turno'] != 'Integral', 'Turno'] = 'Regular'
dados_excel['Data de desistência'] = pd.to_datetime(dados_excel['Data de desistência'])
dados_excel['Ano_desistencia'] = dados_excel['Data de desistência'].dt.year
dados_excel['Mês'] = dados_excel['Data de desistência'].dt.month


########################################### ALUNOS TOTAL ################################

caminho_arquivo_alunos_total = r'G:\.shortcut-targets-by-id\1kArAZwgCxrjbQwQOPEzeJLtMUll3VVJ7\13. Dados\13.2. RMR\13.2.2. Escolas\13.2.2.1. UNI\13.2.2.1.2. Quantitativo\13.2.2.1.2.1. Base\13.2.2.1.2.1.2. Base Escolas\11 - Listagem_dos_alunos.xlsx'
nome_aba_alunos_total = '15 - Listagem dos alunos - atua'
dados_excel_alunos_total = pd.read_excel(caminho_arquivo_alunos_total, sheet_name=nome_aba_alunos_total)

############################################# Tabela desistencia Integral
############################# INDICADOR : nº matrículas turno integral que desistiram (sem ser por bloqueamento por inadimplência) / nº alunos matriculados
# Criar a lista de meses de 1 a 12
meses = list(range(1, 13))
ano_atual = datetime.now().year

# Carregar o arquivo de métrica
planilha = openpyxl.load_workbook(r'G:\.shortcut-targets-by-id\1kArAZwgCxrjbQwQOPEzeJLtMUll3VVJ7\13. Dados\13.2. RMR\13.2.2. Escolas\13.2.2.1. UNI\13.2.2.1.3. Métricas\1. Novos Indicadores Escola.xlsx')

# Selecionar a planilha desejada
Aba_metricas = planilha['Planilha1']

# Puxar o valor da célula A1
Valor_frente1_desistencia_integral = Aba_metricas['A5'].value
Valor_area_responsavel_desistencia_integral = Aba_metricas['B5'].value
Valor_setor_responsavel_desistencia_integral = Aba_metricas['C5'].value
Valor_frente2_desistencia_integral = Aba_metricas['D5'].value
Valor_indicador_desistencia_integral = Aba_metricas['E5'].value
Valor_meta_desistencia_integral = Aba_metricas['F5'].value
Valor_ruim_desistencia_integral = Aba_metricas['G5'].value
Valor_regular_desistencia_integral = Aba_metricas['H5'].value
Valor_otimo_desistencia_integral = Aba_metricas['I5'].value
Valor_peso_desistencia_integral = Aba_metricas['J5'].value


# Criar os dados para a tabela
dados = {
    'Ano': [ano_atual] * len(meses),
    'Mês': meses,
    'Unidade': ["UNI"] * len(meses),
    'FRENTE 1': Valor_frente1_desistencia_integral,
    'ÁREA RESPONSÁVEL RESULTADO': Valor_area_responsavel_desistencia_integral,
    'SETOR RESPONSÁVEL COLETA DADO': Valor_setor_responsavel_desistencia_integral,
    'FRENTE 2': Valor_frente2_desistencia_integral,
    'INDICADOR': Valor_indicador_desistencia_integral,
    'META': Valor_meta_desistencia_integral,
    'RUIM': Valor_ruim_desistencia_integral,
    'REGULAR': Valor_regular_desistencia_integral,
    'ÓTIMO': Valor_otimo_desistencia_integral,
    'Valor': [0] * len(meses),
    'Tipo':["Porcentagem"] * len(meses),
    'Peso' : Valor_peso_desistencia_integral,
}

# Criar o DataFrame
df = pd.DataFrame(dados)

# Inicializar as listas
Desistencia_Integral = []

# Calcular o total de alunos na disciplina
Qtde_total_alunos_disciplina = dados_excel_alunos_total['Numero de Matrícula'].count()

if 'Integral' in dados_excel['Turno'].values:
    Desistencia_Integral = []
    for mes in meses:
        if dados_excel['Mês'].isna().all() or Qtde_total_alunos_disciplina == 0:
            Desistencia_Integral.append(None)
        else:
            # Filtrar os dados para o mês atual e o ano atual
            dados_mes_atual = dados_excel[(dados_excel['Mês'] == mes) & (dados_excel['Ano_desistencia'] == ano_atual)]

            # Calcular a quantidade de desistentes do turno Regular no mês atual
            qtde_Desistente_Integral_total = dados_mes_atual[dados_mes_atual['Turno'] == 'Integral'].shape[0]
            medidas_todo_mes_integral = dados_mes_atual[dados_mes_atual['Segmento'] == 'TODOS'].count()

            # Tem o campo de "não houve desistente" informado
            if  (medidas_todo_mes_integral == 1).any():
                # Calcular a taxa de desistência regular e adicionar à lista
                if Qtde_total_alunos_disciplina != 0:
                    taxa_desistencia_integral = qtde_Desistente_Integral_total / Qtde_total_alunos_disciplina
                else:
                    taxa_desistencia = 0
                Desistencia_Integral.append(taxa_desistencia_integral)
            else:
                Desistencia_Integral.append(0)
else:
    # Se nenhum valor 'Integral' estiver presente na coluna 'Turno'
    Desistencia_Integral = [None] * len(meses)

# Se você já tem a lista Desistencia_Integral, pode convertê-la para uma Series do Pandas
Desistencia_Integral_series = pd.Series(Desistencia_Integral)

# Substituir valores None por 'N/A' na série usando fillna
Desistencia_Integral_series = Desistencia_Integral_series.fillna('N/A')

# Formatar os valores como porcentagens e atribuir ao DataFrame
df['Valor'] = Desistencia_Integral_series.map(lambda x: '{:.2%}'.format(x) if x != 'N/A' else x)


Tratamento_Desistencia_Integral = r'G:\.shortcut-targets-by-id\1kArAZwgCxrjbQwQOPEzeJLtMUll3VVJ7\13. Dados\13.2. RMR\13.2.2. Escolas\13.2.2.1. UNI\13.2.2.1.2. Quantitativo\13.2.2.1.2.1. Base\13.2.2.1.2.1.2. Base Escolas\13.2.2.1.2.1.2.3. Excel\Tratamento_Desistencia_Integral.xlsx'
df.to_excel(Tratamento_Desistencia_Integral,index=False)
