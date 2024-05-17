import csv

def transferir_csv():
    arquivo_origem = '/Users/imac/Downloads/2024-05-17 - 2024-05-17.csv.csv'
    arquivo_destino = '/Users/imac/Desktop/Esquema Inteligencia/BI-Mercado/csv/Coleta1.csv'
    
    # Lê os dados da planilha de origem
    with open(arquivo_origem, 'r') as f_origem:
        dados_origem = [linha for linha in csv.reader(f_origem)]
    
    # Lê os dados da planilha de destino apenas para contar o número de linhas existentes
    with open(arquivo_destino, 'r') as f_destino:
        num_linhas_destino = sum(1 for _ in csv.reader(f_destino))
    
    # Define o número máximo de linhas a serem transferidas
    maximo_linhas = min(50000, len(dados_origem))
    
    # Adiciona os dados da planilha de origem à planilha de destino
    with open(arquivo_destino, 'a', newline='') as f_destino:
        escritor_destino = csv.writer(f_destino)
        escritor_destino.writerows(dados_origem[:maximo_linhas])
    
    print(f"A transferência de {maximo_linhas} linhas foi concluída com sucesso! O arquivo destino agora tem {num_linhas_destino + maximo_linhas} linhas.")

# Chama a função para fazer a transferência
if __name__ == '__main__':
    transferir_csv()