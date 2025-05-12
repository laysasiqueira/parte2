import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options

# ===== ENTRADA DO USUÁRIO =====
palavra_chave = input("🔍 Digite a palavra-chave para buscar no assunto dos e-mails: ")

# ===== CONFIGURAÇÕES DE DOWNLOAD =====
PASTA_DOWNLOADS = os.path.abspath("downloads")
os.makedirs(PASTA_DOWNLOADS, exist_ok=True)

chrome_options = Options()
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": PASTA_DOWNLOADS,
    "download.prompt_for_download": False,
    "plugins.always_open_pdf_externally": True,
})

# ===== INICIALIZA O CHROME =====
driver = webdriver.Chrome(options=chrome_options)
driver.get("https://mail.google.com/")

input("🔐 Faça login manualmente e pressione ENTER para continuar...")

# ===== BUSCA PELO TERMO DIGITADO =====
search_box = driver.find_element(By.NAME, "q")
search_box.send_keys(f"subject:{palavra_chave}")
search_box.send_keys(Keys.RETURN)
time.sleep(5)

# ===== ABRE OS E-MAILS ENCONTRADOS =====
emails = driver.find_elements(By.CSS_SELECTOR, "tr.zA")
print(f"🔍 Encontrados {len(emails)} e-mails com '{palavra_chave}' no assunto.")

for i, email_row in enumerate(emails[:5], start=1):
    print(f"\n📨 Abrindo e-mail {i}")
    email_row.click()
    time.sleep(4)

    # Tenta baixar anexos
    try:
        anexos = driver.find_elements(By.CSS_SELECTOR, "div.aQH span.aZo, div.aQH div.aQy")
        for j, anexo in enumerate(anexos, start=1):
            print(f"⬇️ Baixando anexo {j}")
            anexo.click()
            time.sleep(2)
    except Exception as e:
        print("⚠️ Nenhum anexo encontrado ou erro:", e)

    # Volta para a lista de e-mails
    driver.back()
    time.sleep(4)

print(f"\n✅ Finalizado. Anexos salvos em: {PASTA_DOWNLOADS}")
# driver.quit()
