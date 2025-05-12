import os
import pdfplumber
import pandas as pd
from config import PASTA_DESTINO, LOG_PATH
from datetime import datetime

def log(mensagem):
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(f"[{datetime.now()}] {mensagem}\n")

def main():
    for arquivo in os.listdir(PASTA_DESTINO):
        if arquivo.lower().endswith(".pdf"):
            caminho_pdf = os.path.join(PASTA_DESTINO, arquivo)
            nome_excel = arquivo.replace(".pdf", ".xlsx")
            caminho_excel = os.path.join(PASTA_DESTINO, nome_excel)

            try:
                with pdfplumber.open(caminho_pdf) as pdf:
                    todas_tabelas = []
                    for pagina in pdf.pages:
                        tabela = pagina.extract_table()
                        if tabela:
                            df = pd.DataFrame(tabela[1:], columns=tabela[0])
                            todas_tabelas.append(df)
                    if todas_tabelas:
                        resultado = pd.concat(todas_tabelas, ignore_index=True)
                        resultado.to_excel(caminho_excel, index=False)
                        log(f"PDF convertido para Excel: {caminho_excel}")
                    else:
                        log(f"Nenhuma tabela encontrada em {arquivo}")
            except Exception as e:
                log(f"Erro ao processar {arquivo}: {str(e)}")