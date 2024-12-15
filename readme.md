# Projeto de Extração de Tabelas de PDF para Excel

Este projeto demonstra como extrair dados de tabelas em um arquivo PDF e salvar essas informações em um arquivo Excel utilizando Python. Ele usa as bibliotecas `pdfplumber` para extrair as tabelas do PDF e `pandas` para criar a planilha Excel.

---

## Funcionalidades Principais

1. **Extração de Tabelas de PDF:**
    - O programa abre um arquivo PDF e extrai todas as tabelas presentes em cada página.
    - As tabelas extraídas são armazenadas em uma lista para posterior manipulação.

2. **Salvar Dados em Excel:**
    - Após extrair as tabelas, os dados são convertidos para um formato que pode ser facilmente manipulado em uma planilha.
    - O arquivo Excel é gerado com os dados extraídos, sendo possível visualizá-los e editá-los em qualquer editor de planilhas (como o Microsoft Excel ou o Google Sheets).

---

## Estrutura do Código

### Funções

- **`extrair_tabelas_pdf(pdf_path)`**
    - Esta função recebe o caminho de um arquivo PDF, abre o PDF e extrai todas as tabelas contidas nas páginas.
    - Retorna uma lista com os dados das tabelas extraídas.

- **`salvar_em_excel(dados, output_excel)`**
    - Esta função recebe os dados extraídos do PDF e salva-os em um arquivo Excel.
    - Utiliza a biblioteca `pandas` para gerar a planilha, que é salva no caminho especificado.

---

## Fluxo do Programa

1. **Leitura do PDF:**
    - O programa abre o arquivo PDF fornecido.
    - Cada página do PDF é analisada em busca de tabelas, as quais são extraídas.

2. **Extração de Dados das Tabelas:**
    - Cada tabela é dividida em linhas e colunas, e os dados são organizados de forma a serem armazenados em uma lista.

3. **Criação do Arquivo Excel:**
    - Após a extração, os dados são passados para o `pandas DataFrame` e, em seguida, salvos em um arquivo Excel no formato `.xlsx`.

4. **Exibição do Resultado:**
    - O programa informa que a extração foi realizada com sucesso e que os dados foram salvos no arquivo Excel especificado.

---

## Como Executar o Projeto

1. **Clone o repositório.**

2. **Configure o ambiente Python.**
    - Certifique-se de ter o **Python 3.x** instalado.
    - Instale as bibliotecas necessárias utilizando o comando:
      ```bash
      pip install pdfplumber pandas openpyxl
      ```

3. **Execute o código Python.**
    - Execute o arquivo Python (`extrair_pdf_para_excel.py` ou similar):
      ```bash
      python extrair_pdf_para_excel.py
      ```

4. **Veja os resultados.**
    - O programa criará um arquivo Excel com os dados extraídos e salvará o arquivo com o nome especificado no código (por exemplo, `saida.xlsx`).

---

## Exemplo de Uso

1. **Input:**
    - Um arquivo PDF (`arquivo.pdf`) contendo tabelas.

2. **Output:**
    - Um arquivo Excel (`saida.xlsx`) com as tabelas extraídas do PDF.

---

## Tecnologias Utilizadas

- **Python 3.x:** Linguagem de programação utilizada.
- **pdfplumber:** Biblioteca para extração de dados de arquivos PDF.
- **pandas:** Biblioteca para manipulação de dados e criação de planilhas Excel.
- **openpyxl:** Biblioteca necessária para salvar o arquivo no formato Excel `.xlsx`.

---

## Autor
Lucas Degrande