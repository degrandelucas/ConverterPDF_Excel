import pdfplumber
import pandas as pd

def extrair_tabelas_pdf(pdf_path):
    dados_tabela = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for pagina in pdf.pages:
            tabelas = pagina.extract_tables()
            for tabela in tabelas:
                for linha in tabela:
                    dados_tabela.append(linha)  #Adiciona cada linha da tabela
    
    return dados_tabela

def salvar_em_excel(dados, output_excel):
    df = pd.DataFrame(dados)
    df.to_excel(output_excel, index=False, header=False)

pdf_path = 'arquivo.pdf'
output_excel = 'saida.xlsx'

dados_extraidos = extrair_tabelas_pdf(pdf_path)
salvar_em_excel(dados_extraidos, output_excel)

print(f"Os dados da tabela foram extraídos e salvos em {output_excel}")