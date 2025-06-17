# Projeto de Análise Comparativa de Ordens de Serviço (OS)

Este projeto utiliza Python e Pandas para realizar uma análise comparativa de Ordens de Serviço (OS) entre dados extraídos de um sistema BI e dados de um sistema de Produção.

## Funcionalidades

1.  **Seleção de OS do BI:**
    *   Carrega uma planilha Excel (`BI.xlsx`).
    *   Seleciona as 3 OS com maior "Fat Total", as 3 com menor "Fat Total" e 4 aleatórias.
    *   Gera um relatório HTML (`OS_selecionadas_BI_completo.html`) com a lista, soma total e um gráfico de barras do "Fat Total".
    *   Gera um arquivo Excel (`OS_selecionadas_BI.xlsx`) com as OS selecionadas.
    *   Script: `step1_analise_bi.py` (ou o nome que você usou para o primeiro script)

2.  **Comparação BI vs. Produção:**
    *   Carrega as OS selecionadas do BI (`OS_selecionadas_BI.xlsx`).
    *   Carrega uma planilha de dados da Produção (`planilha_producao_teste.xlsx`).
    *   Compara o "Fat Total" para cada OS.
    *   Gera um relatório HTML (`relatorio_comparativo_BI_x_Producao.html`) com a tabela comparativa e destaques visuais para as diferenças.
    *   Script: `comparar_os_html.py` (ou o nome que você usou para o segundo script)

## Pré-requisitos

*   Python 3.x
*   Bibliotecas Python:
    *   pandas
    *   openpyxl
    *   matplotlib
    *   seaborn

## Instalação

1.  Clone o repositório:
    ```bash
    git clone https://github.com/SEU_USUARIO/NOME_DO_SEU_REPOSITORIO.git
    cd NOME_DO_SEU_REPOSITORIO
    ```

2.  (Recomendado) Crie e ative um ambiente virtual:
    ```bash
    python -m venv venv
    # Windows
    venv\Scripts\activate
    # macOS/Linux
    source venv/bin/activate
    ```

3.  Instale as dependências (crie um arquivo `requirements.txt` primeiro - veja abaixo):
    ```bash
    pip install -r requirements.txt
    ```

## Como Usar

1.  Coloque seus arquivos de dados `BI.xlsx` e `planilha_producao_teste.xlsx` na pasta raiz do projeto (ou ajuste os nomes dos arquivos nos scripts).
2.  Execute o script de análise do BI:
    ```bash
    python step1_analise_bi.py
    ```
3.  Execute o script de comparação:
    ```bash
    python comparar_os_html.py
    ```
4.  Abra os arquivos `.html` gerados no seu navegador para ver os resultados.

## Arquivos Principais

*   `step1_analise_bi.py`: Script para selecionar e analisar OS do BI.
*   `comparar_os_html.py`: Script para comparar OS do BI com Produção.
*   `BI.xlsx`: (Exemplo) Planilha de entrada com dados do BI.
*   `planilha_producao_teste.xlsx`: (Exemplo) Planilha de entrada com dados da produção.
*   `.gitignore`: Especifica arquivos intencionalmente não rastreados.
*   `README.md`: Este arquivo.
*   `requirements.txt`: Lista de dependências Python.
