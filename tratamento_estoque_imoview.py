import pandas as pd

# Carregar os dados da planilha de imóveis
imoveis_df = pd.read_excel("/Users/imac/Downloads/imoveis-2024-05-17-095150.xlsx")

# Garantindo que todas as colunas usadas na concatenação sejam do tipo string
imoveis_df['Endereco'] = imoveis_df['Endereco'].astype(str)
imoveis_df['Bloco'] = imoveis_df.get('Bloco', '').astype(str)  # Usa coluna 'Bloco' se existir, senão cria vazia
imoveis_df['Complemento'] = imoveis_df.get('Complemento', '').astype(str)  # Usa 'Complemento' se existir, senão cria vazia

# Concatenando as informações de endereço
imoveis_df['Endereco'] = imoveis_df['Endereco'] + " " + imoveis_df['Bloco'] + " " + imoveis_df['Complemento']

# Selecionando e renomeando colunas conforme a estrutura de saída desejada
# Ordenando as colunas conforme a especificação
transformed_df = imoveis_df[['Codigo', 'Captadores', 'Tipo', 'NumeroQuarto', 'Valor', 'Endereco', 'Bairro']].copy()
transformed_df['Captadores'] = imoveis_df.get('Captadores', 'N/D')  # Usa 'Captadores' se existir, senão 'N/D'
transformed_df['NumeroQuarto'] = imoveis_df.get('NumeroQuarto', 'N/D')  # Usa 'NumeroQuarto' se existir, senão 'N/D'
transformed_df['Bairro'] = imoveis_df.get('Bairro', 'N/D')  # Usa 'Bairro' se existir, senão 'N/D'

# Salvar a planilha transformada no formato desejado
transformed_df.to_excel("/Users/imac/Downloads/Estoque 17-05-2024.xlsx", index=False)

# Imprimindo as primeiras linhas para conferência
print(transformed_df.head())
