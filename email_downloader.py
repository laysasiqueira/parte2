# RPA/email_downloader.py

import imaplib
import email
from email.header import decode_header
import os
from datetime import datetime
from config import EMAIL, SENHA, IMAP_SERVER, PASTA_DESTINO, LOG_PATH

def log(mensagem):
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(f"[{datetime.now()}] {mensagem}\n")

def conectar_email():
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL, SENHA)
    mail.select("inbox")
    return mail

def buscar_emails(mail):
    status, mensagens = mail.search(None, '(UNSEEN)')
    return mensagens[0].split()

def baixar_anexos(mail):
    os.makedirs(PASTA_DESTINO, exist_ok=True)
    os.makedirs(os.path.dirname(LOG_PATH), exist_ok=True)

    mensagens = buscar_emails(mail)

    with open(LOG_PATH, "a", encoding="utf-8") as log:
        for num in mensagens:
            status, dados = mail.fetch(num, '(RFC822)')
            raw_email = dados[0][1]
            msg = email.message_from_bytes(raw_email)

            assunto = msg.get("subject", "Sem assunto")
            remetente = msg.get("from", "Desconhecido")

            for parte in msg.walk():
                if parte.get_content_maintype() == 'multipart':
                    continue
                if parte.get("Content-Disposition") is None:
                    continue

                nome_arquivo = parte.get_filename()
                if nome_arquivo:
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    caminho = os.path.join(PASTA_DESTINO, f"{timestamp}_{nome_arquivo}")

                    with open(caminho, "wb") as f:
                        f.write(parte.get_payload(decode=True))

                    log.write(f"{datetime.now()} - Anexo salvo: {caminho} | De: {remetente} | Assunto: {assunto}\n")
                    print(f"📎 Anexo salvo: {caminho}")
