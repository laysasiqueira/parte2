# RPA/main.py

import os
from email_downloader import conectar_email, buscar_emails, baixar_anexos
from config import PASTA_DESTINO, LOG_PATH

# Criar pastas se não existirem
os.makedirs(PASTA_DESTINO, exist_ok=True)
os.makedirs(os.path.dirname(LOG_PATH), exist_ok=True)

def executar_rpa():
    try:
        palavra_chave = input("🔍 Digite a palavra-chave para buscar no assunto: ").strip()
        if not palavra_chave:
            raise ValueError("A palavra-chave não pode estar vazia.")

        mail = conectar_email()
        mensagens = buscar_emails(mail, palavra_chave)
        baixar_anexos(mail, mensagens)
        print("✅ RPA concluído com sucesso.")
    except Exception as e:
        print(f"❌ Erro: {e}")

def new_func():
    return True

if __name__ == "__main__":
    executar_rpa()
