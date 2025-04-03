# sheet-folder-organizer
Organizador de Arquivos com Planilha Excel
Este projeto automatiza a organização de arquivos e pastas com base nos dados de uma planilha do Excel. Ele verifica uma condição específica e move/copía as pastas correspondentes para um diretório de destino, organizando-as conforme um padrão definido.

# 🚀 Funcionalidades
✅ Lê uma planilha Excel (.xlsx) e identifica clientes com base em uma condição definida pelo usuário.
✅ Verifica a existência das pastas dos clientes no diretório de origem.
✅ Copia os arquivos das pastas encontradas para um diretório de destino, estruturando-os com base em mês e ano.
✅ Gera relatórios (.txt) listando os clientes transferidos e aqueles cuja pasta não foi encontrada.

📦 Tecnologias Utilizadas
Python 🐍

Pandas 📊 (para manipulação de planilhas Excel)

OS & Shutil 📂 (para manipulação de arquivos e diretórios)

⚙️ Como Usar
1️⃣ Baixe ou clone o repositório:

bash
Copiar
Editar
git clone https://github.com/seu-usuario/seu-repositorio.git
cd seu-repositorio
2️⃣ Instale as dependências necessárias:

bash
Copiar
Editar
pip install pandas openpyxl
3️⃣ Execute o script e siga as instruções:

bash
Copiar
Editar
python organizador.py
4️⃣ Insira os dados solicitados:

Caminho do arquivo Excel

Colunas que contêm o número do cliente e a condição desejada

Diretórios de origem e destino

Mês e ano no formato MMAAAA (exemplo: 022024 para fevereiro de 2024)

5️⃣ O programa moverá/copiará os arquivos conforme a condição filtrada.

#📑 Exemplo de Uso
📊 Planilha de entrada (clientes.xlsx)
A (Cliente)	B (Status)
1001	Aprovado
1002	Rejeitado
1003	Aprovado
🎯 Configuração no Script
Coluna do número do cliente: A

Coluna da condição: B

Condição desejada: "Aprovado"

📂 O script moverá as pastas correspondentes aos clientes 1001 e 1003 para o diretório de destino, criando uma subpasta no formato Cliente-MMAAAA.

📜 Relatórios Gerados
📌 clientes_sem_pasta.txt → Lista de clientes cuja pasta não foi encontrada no diretório de origem.
📌 clientes_transferidos.txt → Lista de clientes cujos arquivos foram copiados com sucesso.

📌 Melhorias Futuras
🔹 Adicionar interface gráfica para facilitar o uso.
🔹 Permitir mover arquivos em vez de copiar.
🔹 Suporte para múltiplas condições de filtragem.

📝 Licença
Este projeto está licenciado sob a licença MIT.

