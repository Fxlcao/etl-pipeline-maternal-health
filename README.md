# 🤰 Maternal Health Data Pipeline - ETL & Automation

[![Python](https://img.shields.io/badge/Python-3.10+-blue.svg)](https://www.python.org/)
[![Pandas](https://img.shields.io/badge/Library-Pandas-150458.svg)](https://pandas.pydata.org/)
[![Status](https://img.shields.io/badge/Status-Project_Ready-green.svg)]()

## 📌 Contexto e Objetivo
Este projeto foi desenvolvido para resolver um desafio comum na área de saúde: a fragmentação de dados. O script automatiza o processo de **ETL (Extract, Transform, Load)** de informações do projeto **AGAR**, consolidando dados de avaliações de risco gestacional (Estratificação), triagem de saúde mental (EPDS) e consultas de enfermagem.

A ferramenta transforma planilhas de Excel complexas e inconsistentes em um **Dataset consolidado e higienizado (CSV)**, modelado especificamente para alimentar dashboards de alta performance no **Power BI**.

## 🛠️ Stack Tecnológica
* **Python**: Linguagem base para o processamento.
* **Pandas**: Biblioteca principal para manipulação de grandes volumes de dados e lógica de junção.
* **Tkinter**: Implementação de interface gráfica (GUI) para facilitar o uso por usuários não técnicos.
* **OpenPyXL**: Engine para leitura e escrita de arquivos Excel.

## ⚙️ Diferenciais Técnicos (Data Wrangling)
O diferencial deste projeto não é apenas a junção das tabelas, mas a inteligência aplicada no tratamento dos dados:

1.  **Lógica de Junção (Left Join)**: O script utiliza a aba 'Cadastro EPDS' como tabela fato principal, realizando o cruzamento com as demais abas via "Nome Social", garantindo a integridade dos registros.
2.  **Limpeza via RegEx**: Identificação e normalização de strings vazias ou preenchidas apenas com espaços em branco, padronizando a entrada para o banco de dados.
3.  **Parsing de Strings**: Extração inteligente da pontuação de estratificação (removendo decimais desnecessários e caracteres de formatação).
4.  **Imputação Condicional de Nulos**: 
    * Colunas **numéricas/métricas** sem dados são preenchidas com `"Null"`.
    * Colunas **textuais/dimensões** recebem o marcador `"-"`.
    * Isso evita erros de tipagem e "sujeira" visual no Power BI (como o termo "Blank").
5.  **Interface de Usuário**: O fluxo de seleção de arquivos e salvamento é totalmente interativo, eliminando a necessidade de interação com o terminal.

## 📂 Como Utilizar
1. Instale as dependências necessárias:
   ```bash
   pip install pandas openpyxl

### ⚠️ Aviso de Privacidade (LGPD)
Por motivos de segurança da informação e conformidade com a LGPD, este repositório contém **estritamente a lógica do script**, sem qualquer base de dados associada. O foco aqui é a demonstração da arquitetura de ETL e automação.
