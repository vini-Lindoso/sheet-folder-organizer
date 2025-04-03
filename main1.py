import os
import shutil
import pandas as pd

# Função para criar o diretório de destino (se não existir)
def criar_diretorio_se_nao_existir(caminho):
    if not os.path.exists(caminho):
        os.makedirs(caminho)
        print(f"Diretório criado: {caminho}")
    else:
        print(f"Diretório já existe: {caminho}")

# Solicita o caminho da planilha Excel
planilha_excel = input('Cole o caminho da planilha Excel: ')

# Verifica se o arquivo existe
if not os.path.exists(planilha_excel):
    print(f"Erro: O arquivo '{planilha_excel}' não foi encontrado.")
    exit()

# Lê o arquivo Excel sem considerar a primeira linha como cabeçalho
df = pd.read_excel(planilha_excel, header=None)

# Verifica o número de colunas no DataFrame
num_colunas = df.shape[1]

# Define os nomes das colunas com base no número de colunas
nomes_colunas = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'][:num_colunas]
df.columns = nomes_colunas

# Exibe as primeiras linhas do DataFrame para verificar
print(df.head())

# Solicita os nomes das colunas
coluna_numero_cliente = input('Digite o nome da coluna que contém o número do cliente: ')
coluna_condicao = input('Digite o nome da coluna que contém a condição: ')
condicao = input('Digite a condição que deseja filtrar: ')

# Verifica se as colunas existem no DataFrame
if coluna_numero_cliente not in df.columns or coluna_condicao not in df.columns:
    print(f"Erro: Uma das colunas '{coluna_numero_cliente}' ou '{coluna_condicao}' não existe na planilha.")
    print(f"Colunas disponíveis: {list(df.columns)}")
    exit()

# Filtra os clientes com a condição especificada
clientes_para_mover = df.loc[df[coluna_condicao] == condicao, coluna_numero_cliente]

# Verifica se algum cliente foi encontrado
if clientes_para_mover.empty:
    print(f"Erro: Nenhum cliente encontrado com a condição '{condicao}'.")
    exit()

# Solicita o caminho da pasta de origem (onde estão as pastas dos clientes)
pasta_origem = input('Digite o caminho da pasta de origem (onde estão as pastas dos clientes): ')

# Solicita o caminho da pasta de destino (onde as pastas serão copiadas)
pasta_destino = input('Digite o caminho da pasta de destino: ')

# Solicita o mês e ano para criar a pasta no formato MMAAAA
mes_ano = input('Digite o mês e ano no formato MMAAAA (ex: 022024 para fevereiro de 2024): ')

# Lista para armazenar clientes sem pasta na origem
clientes_sem_pasta = []

# Lista para armazenar clientes transferidos com sucesso
clientes_transferidos = []

# Itera sobre os clientes com a condição especificada
for cliente in clientes_para_mover:
    # Caminho da pasta do cliente na origem
    pasta_cliente_origem = os.path.join(pasta_origem, str(cliente))
    
    # Verifica se a pasta do cliente existe na origem
    if not os.path.exists(pasta_cliente_origem):
        print(f"Pasta do cliente {cliente} não encontrada em {pasta_origem}.")
        clientes_sem_pasta.append(cliente)  # Adiciona à lista de clientes sem pasta
        continue

    # Caminho da pasta de destino (pasta do cliente + pasta MMAAAA)
    pasta_cliente_destino = os.path.join(pasta_destino, f"{cliente}-", mes_ano)
    
    # Cria o diretório de destino (se não existir)
    criar_diretorio_se_nao_existir(pasta_cliente_destino)

    # Copia os arquivos da pasta de origem para a pasta de destino
    for arquivo in os.listdir(pasta_cliente_origem):
        caminho_arquivo_origem = os.path.join(pasta_cliente_origem, arquivo)
        caminho_arquivo_destino = os.path.join(pasta_cliente_destino, arquivo)
        
        # Copia o arquivo
        shutil.copy2(caminho_arquivo_origem, caminho_arquivo_destino)
        print(f"Arquivo '{arquivo}' copiado para '{pasta_cliente_destino}'.")

    # Adiciona o cliente à lista de transferidos
    clientes_transferidos.append(cliente)

# Gera um arquivo TXT com a lista de clientes sem pasta na origem
if clientes_sem_pasta:
    caminho_txt_sem_pasta = os.path.join(pasta_destino, 'clientes_sem_pasta.txt')
    with open(caminho_txt_sem_pasta, 'w') as arquivo_txt:
        for cliente in clientes_sem_pasta:
            arquivo_txt.write(f"{cliente}\n")
    print(f"Arquivo 'clientes_sem_pasta.txt' gerado em: {caminho_txt_sem_pasta}")
else:
    print("Todos os clientes possuem pastas na origem.")

# Gera um arquivo TXT com a lista de clientes transferidos com sucesso
if clientes_transferidos:
    caminho_txt_transferidos = os.path.join(pasta_destino, 'clientes_transferidos.txt')
    with open(caminho_txt_transferidos, 'w') as arquivo_txt:
        for cliente in clientes_transferidos:
            arquivo_txt.write(f"{cliente}\n")
    print(f"Arquivo 'clientes_transferidos.txt' gerado em: {caminho_txt_transferidos}")
else:
    print("Nenhum cliente foi transferido.")

print("Processo concluído!")