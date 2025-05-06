# RPA/email_downloader.py

import imaplib
import email
from email.header import decode_header
from email.utils import parseaddr
import os
from datetime import datetime
from openpyxl import load_workbook
from config import EMAIL, SENHA, IMAP_SERVER, PASTA_DESTINO, LOG_PATH, WHITELISTED_SENDERS

# Extensões suportadas
EXTENSOES_PERMITIDAS = ('.xlsx', '.xls', '.pdf', '.png', '.jpg', '.jpeg', '.docx', '.doc', '.txt')

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

def abrir_excel(path):
    try:
        print(f"📊 Abrindo e lendo Excel: {path}")
        wb = load_workbook(filename=path)
        sheet = wb.active
        for row in sheet.iter_rows(values_only=True):
            print(row)
    except Exception as e:
        print(f"⚠️ Erro ao ler Excel {path}: {e}")

def baixar_anexos(mail):
    os.makedirs(PASTA_DESTINO, exist_ok=True)
    os.makedirs(os.path.dirname(LOG_PATH), exist_ok=True)

    mensagens = buscar_emails(mail)

    with open(LOG_PATH, "a", encoding="utf-8") as log_file:
        for num in mensagens:
            status, dados = mail.fetch(num, '(RFC822)')
            raw_email = dados[0][1]
            msg = email.message_from_bytes(raw_email)

            assunto = msg.get("subject", "Sem assunto")
            remetente = msg.get("from", "Desconhecido")
            remetente_email = parseaddr(remetente)[1]

            if remetente_email not in WHITELISTED_SENDERS:
                print(f"⛔ Remetente não autorizado: {remetente_email}")
                log_file.write(f"{datetime.now()} - Email ignorado de: {remetente_email}\n")
                continue

            for parte in msg.walk():
                if parte.get_content_maintype() == 'multipart':
                    continue
                if parte.get("Content-Disposition") is None:
                    continue

                nome_arquivo = parte.get_filename()
                if not nome_arquivo:
                    continue

                # Garante que a extensão seja permitida
                if not nome_arquivo.lower().endswith(EXTENSOES_PERMITIDAS):
                    print(f"⚠️ Tipo de arquivo ignorado: {nome_arquivo}")
                    continue

                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                caminho = os.path.join(PASTA_DESTINO, f"{timestamp}_{nome_arquivo}")

                with open(caminho, "wb") as f:
                    f.write(parte.get_payload(decode=True))

                log_file.write(f"{datetime.now()} - Anexo salvo: {caminho} | De: {remetente_email} | Assunto: {assunto}\n")
                print(f"📎 Anexo salvo: {caminho}")

                # Se for Excel, abre o conteúdo
                if nome_arquivo.endswith(".xlsx"):
                    abrir_excel(caminho)
