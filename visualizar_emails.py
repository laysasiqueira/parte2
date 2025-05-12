# visualizar_emails.py

from selenium import webdriver
from selenium.webdriver.common.by import By
import time

def visualizar_emails_no_gmail(palavra_chave):
    print(f"🌐 Abrindo navegador para mostrar e-mails com '{palavra_chave}' no Gmail...")

    # Abrir o navegador Chrome
    driver = webdriver.Chrome()
    driver.get("https://mail.google.com")

    # Espera o usuário fazer login
    input("🔐 Faça login no Gmail no navegador que abriu e pressione ENTER aqui para continuar...")

    time.sleep(5)

    # Buscar pela palavra-chave no campo de busca
    search_box = driver.find_element(By.NAME, "q")
    search_box.clear()
    search_box.send_keys(f"subject:{palavra_chave}")
    search_box.submit()

    print("📬 Busca enviada, veja o Gmail no navegador.")
    time.sleep(15)  # Tempo para o usuário ver os resultados

    # driver.quit()  # Descomente se quiser fechar o navegador automaticamente
