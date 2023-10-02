import os
import pandas as pd

pastas = [
    r'G:\.shortcut-targets-by-id\1kArAZwgCxrjbQwQOPEzeJLtMUll3VVJ7\13. Dados\2 - RMR\2.3 - ESCOLAS\2.3.2 - MVM\2.3.2 - QUANTITATIVO\2.3.2.1 - BASE\2.3.2.1.2 - Base Escolas\Excel',
    r'G:\.shortcut-targets-by-id\1kArAZwgCxrjbQwQOPEzeJLtMUll3VVJ7\13. Dados\2 - RMR\2.3 - ESCOLAS\2.3.1 - UNI\2.3.1.2 - QUANTITATIVO\2.3.2.1 - BASE\2.3.2.1.2 - Base Escolas\Excel'
]
dados_concatenados = pd.DataFrame()
# Loop para percorrer os arquivos na pasta
for pasta in pastas:
    for arquivo in os.listdir(pasta):
        caminho_arquivo = os.path.join(pasta, arquivo)
        if os.path.isfile(caminho_arquivo) and arquivo.endswith('.xlsx'):
            dados_planilha = pd.read_excel(caminho_arquivo)
            dados_concatenados = pd.concat([dados_concatenados, dados_planilha])

caminho_destino = r'G:\.shortcut-targets-by-id\1kArAZwgCxrjbQwQOPEzeJLtMUll3VVJ7\13. Dados\2 - RMR\2.3 - ESCOLAS\2.3.0 - GERENCIAL\2.3.2 - QUANTITATIVO\2.3.2.1 - BASE\2.3.2.1.0 - Consolidado\0. Gerencial\consolidado.xlsx'  # Substitua pelo caminho desejado
# Salvar os dados concatenados em uma única planilha no arquivo Excel no caminho especificado
dados_concatenados.to_excel(caminho_destino, sheet_name='Dados Consolidados', index=False)
