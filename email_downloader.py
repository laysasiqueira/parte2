# email_downloader.py

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

def baixar_anexos(mail):
    status, mensagens = mail.search(None, "ALL")
    for num in mensagens[0].split():
        status, dados = mail.fetch(num, "(RFC822)")
        raw_email = dados[0][1]
        msg = email.message_from_bytes(raw_email)

        assunto, codif = decode_header(msg["Subject"])[0]
        assunto = assunto.decode(codif or "utf-8") if isinstance(assunto, bytes) else assunto
        de = msg.get("From")
        log(f"E-mail recebido de {de} com assunto '{assunto}'.")

        for parte in msg.walk():
            if parte.get_content_maintype() == "multipart":
                continue
            if parte.get("Content-Disposition") is None:
                continue

            nome_arquivo = parte.get_filename()
            if nome_arquivo:
                nome_arquivo, codif = decode_header(nome_arquivo)[0]
                if isinstance(nome_arquivo, bytes):
                    nome_arquivo = nome_arquivo.decode(codif or "utf-8")

                caminho = os.path.join(PASTA_DESTINO, nome_arquivo)
                with open(caminho, "wb") as f:
                    f.write(parte.get_payload(decode=True))
                log(f"Anexo salvo: {caminho}")

    mail.logout()
