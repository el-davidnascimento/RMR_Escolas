import pandas as pd

# Carregando as planilhas de desistencia
desistencias_universo = pd.read_excel(r'G:/.shortcut-targets-by-id/1kArAZwgCxrjbQwQOPEzeJLtMUll3VVJ7/13. Dados/13.2. RMR/13.2.2. Escolas/13.2.2.1. UNI/13.2.2.1.2. Quantitativo/13.2.2.1.2.1. Base/13.2.2.1.2.1.2. Base Escolas/3-Desistencia_de_alunos_universo.xlsx')
desistencias_messejana = pd.read_excel(r'G:/.shortcut-targets-by-id/1kArAZwgCxrjbQwQOPEzeJLtMUll3VVJ7/13. Dados/13.2. RMR/13.2.2. Escolas/13.2.2.2. MVM/13.2.2.2.2. Quantitativo/13.2.2.2.2.1. Base/13.2.2.2.2.1.2. Base Escolas/3-Desistencia_de_alunos_messejana.xlsx')

# Removendo o campo 'Momento da desistências" e Adicionando o campo 'Unidade' às duas planilhas
desistencias_universo = desistencias_universo.drop('Momento da desistencia', axis=1)
desistencias_universo = desistencias_universo.assign(Unidade="Universo")
desistencias_messejana = desistencias_messejana.drop('Momento da desistencia', axis=1)
desistencias_messejana = desistencias_messejana.assign(Unidade="Messejana")

# Unindo as duas planilhas usando o método concat
planilha_consolidada = pd.concat([desistencias_universo, desistencias_messejana], ignore_index=True)

# Salvando a planilha unida
planilha_consolidada.to_excel(r'G:\.shortcut-targets-by-id/1kArAZwgCxrjbQwQOPEzeJLtMUll3VVJ7/13. Dados/13.2. RMR/13.2.2. Escolas/13.2.2.0. Gerencial/13.2.2.0.2. Quantitativo/13.2.2.0.2.1. Base/13.2.2.0.2.1.1. Base Escolas/3 - Desistencia_de_alunos.xlsx', index=False)