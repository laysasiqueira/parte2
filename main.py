# RPA/main.py

import os
from email_downloader import conectar_email, baixar_anexos
from config import PASTA_DESTINO, LOG_PATH

# Criar pastas se não existirem
os.makedirs(PASTA_DESTINO, exist_ok=True)
os.makedirs(os.path.dirname(LOG_PATH), exist_ok=True)

def executar_rpa():
    try:
        os.makedirs(PASTA_DESTINO, exist_ok=True)
        os.makedirs(os.path.dirname(LOG_PATH), exist_ok=new_func())

        mail = conectar_email()
        baixar_anexos(mail)
        print("✅ RPA concluído com sucesso.")
    except Exception as e:
        print(f"❌ Erro: {e}")

def new_func():
    return True

if __name__ == "__main__":
    executar_rpa()
