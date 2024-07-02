## Descrição Técnica
Este repositório contém um script em Python que utiliza a biblioteca openpyxl para processar e gerenciar dados em arquivos Excel. As principais funcionalidades incluem:

 - Limpeza de Conteúdo: Remove o conteúdo das colunas A a F a partir da linha 2 de um arquivo Excel específico.

 - Cópia de Dados: Copia dados das colunas R, T, V, AF, M, AP de uma planilha de origem para as colunas A, B, C, D, E, F de uma planilha de destino.

 - Preenchimento Automático: Preenche a coluna G da planilha de destino com uma mensagem específica para cada linha preenchida nas colunas A a F.

 - Movimentação de Arquivo: Após o processamento, o arquivo Excel resultante é movido para um diretório específico.

## Estrutura do Repositório
 - script_processamento_excel.py: Contém o código principal do script.
 - LAYOUT AÇÃO DE WHATS TRIGG.xlsx: Arquivo Excel usado como destino para os dados processados.
 - MATRIZ.xlsx: Arquivo Excel usado como fonte de dados para as colunas especificadas.
 - Logs/: Diretório onde são armazenados os arquivos de log.
## Requisitos
 - Python 3.x
 - Biblioteca openpyxl
## Uso
 - Configuração Inicial: Configure os caminhos dos arquivos Excel de origem e destino conforme necessário no script.

 - Execução: Execute o script para processar e gerenciar os dados conforme descrito acima.

### Autor
João Victor da Silva

## Contribuições
Contribuições são bem-vindas! Sinta-se à vontade para enviar pull requests para melhorias ou correções.
