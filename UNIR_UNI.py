import os
import pandas as pd

pasta = r'G:\.shortcut-targets-by-id\1kArAZwgCxrjbQwQOPEzeJLtMUll3VVJ7\13. Dados\13.2. RMR\13.2.2. Escolas\13.2.2.1. UNI\13.2.2.1.2. Quantitativo\13.2.2.1.2.1. Base\13.2.2.1.2.1.2. Base Escolas\13.2.2.1.2.1.2.3. Excel'  # Substitua pelo caminho da pasta desejada
dados_concatenados = pd.DataFrame()
# Loop para percorrer os arquivos na pasta
for arquivo in os.listdir(pasta):
    caminho_arquivo = os.path.join(pasta, arquivo)
    if os.path.isfile(caminho_arquivo) and arquivo.endswith('.xlsx'):
        dados_planilha = pd.read_excel(caminho_arquivo)
        dados_concatenados = pd.concat([dados_concatenados, dados_planilha])

caminho_destino = r'G:\.shortcut-targets-by-id\1kArAZwgCxrjbQwQOPEzeJLtMUll3VVJ7\13. Dados\13.2. RMR\13.2.2. Escolas\13.2.2.0. Gerencial\13.2.2.0.2. Quantitativo\13.2.2.0.2.1. Base\13.2.2.0.2.1.0. Consolidado\13.2.2.0.2.1.0.3. UNI\consolidado.xlsx'  # Substitua pelo caminho desejado
# Salvar os dados concatenados em uma única planilha no arquivo Excel no caminho especificado
dados_concatenados.to_excel(caminho_destino, sheet_name='Dados Consolidados', index=False)
