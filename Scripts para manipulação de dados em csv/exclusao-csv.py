import pandas as pd
from io import StringIO

# Caminho para o seu arquivo CSV
arquivo_csv = '/Users/imac/Desktop/Esquema Inteligencia/BI-Mercado/csv/Coleta1.csv'

# Abre o arquivo em modo de leitura de texto
with open(arquivo_csv, 'r') as file:
    lines = file.readlines()

# Supondo que o separador de colunas seja vírgula
num_colunas_esperado = len(lines[0].split(','))  # Conta o número de colunas na primeira linha (cabeçalho)

# Filtra linhas que têm o número correto de colunas
linhas_corretas = [line for line in lines if len(line.split(',')) == num_colunas_esperado]

# Junta as linhas filtradas de volta em uma string
dados_corrigidos = ''.join(linhas_corretas)

# Converte a string filtrada em um DataFrame do pandas
dados = pd.read_csv(StringIO(dados_corrigidos))

# Coluna de data no seu arquivo CSV
coluna_data = 'Data'

# Converter a coluna de data para o formato de data, transformando datas inválidas em NaT
dados[coluna_data] = pd.to_datetime(dados[coluna_data], errors='coerce')

# Definir o intervalo de datas (inclusive)
data_inicio = pd.to_datetime('2023-07-01')
data_fim = pd.to_datetime('2023-07-01')

# Filtrar dados para excluir linhas que estão dentro do intervalo de datas especificado
dados_filtrados = dados[~((dados[coluna_data] >= data_inicio) & (dados[coluna_data] <= data_fim))]

# Salvar os dados filtrados de volta no arquivo original
dados_filtrados.to_csv(arquivo_csv, index=False)
