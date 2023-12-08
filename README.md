# Relatório Automatizado de Vendas por Loja

Este script Python gera um relatório automatizado de vendas por loja usando a biblioteca pandas para manipulação de dados e a biblioteca win32com.client para enviar o relatório por e-mail utilizando o Outlook.

## Uso

1. Configure a base de dados:
   - Certifique-se de ter um arquivo 'Vendas.xlsx' na mesma pasta do script ou ajuste o caminho do arquivo na linha `tabela_vendas = pd.read_excel('Vendas.xlsx')`.

2. Execute o script:
   - Certifique-se de ter instalado as bibliotecas necessárias do Python. Consulte o arquivo `requirements.txt` para obter detalhes.
   - Execute o script em um ambiente Python.

3. Resultados:
   - O script carrega a base de dados de vendas, realiza análises (faturamento, quantidade de produtos vendidos e ticket médio) por loja e gera um relatório formatado em HTML.
   - O relatório é enviado por e-mail através do Outlook.

## Instalação

1. Clone o repositório:

   ```bash
   git clone https://github.com/seu-usuario/seu-repositorio.git
   cd seu-repositorio
