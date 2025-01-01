from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException
from selenium.webdriver.support import expected_conditions as EC
import time
from openpyxl import Workbook

# Configura o driver
driver = webdriver.Chrome()

try:
    # Acessa o site do Mercado Livre
    driver.get("https://www.mercadolivre.com.br/")
    time.sleep(2)

    # Maximiza a janela
    driver.maximize_window()

    # Identifica a barra de pesquisa e realiza a busca
    # Aguarda o elemento estar visível
    search_box = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.NAME, "as_word"))
    )
    search_box.send_keys("teclado sem fio")
    search_box.send_keys(Keys.RETURN)
    time.sleep(5)

    # Extrai os nomes e preços dos produtos
    product_names = driver.find_elements(By.CSS_SELECTOR, "div.poly-card__content h2.poly-component__title")
    product_prices = driver.find_elements(By.CSS_SELECTOR, "div.poly-price__current")

    # Debugging: Print the number of elements found
    print(f"Found {len(product_names)} product names")
    print(f"Found {len(product_prices)} product prices")

    # Limite de 10 primeiros produtos
    results = []
    for name, price in zip(product_names[:10], product_prices[:10]):
        results.append({"name": name.text, "price": price.text})
        # Debugging: Print each product name and price
        print(f"Product: {name.text}, Price: {price.text}")

    # Salva os dados em um arquivo Excel
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Resultados Mercado Livre"

    # Adiciona os cabeçalhos
    sheet.append(["Nome do Produto", "Preço"])

    # Adiciona os dados
    for result in results:
        sheet.append([result["name"], result["price"]])

    # Salva o arquivo Excel
    workbook.save("mercado_livre_results.xlsx")
    print("Dados salvos em 'mercado_livre_results.xlsx' com sucesso!")

finally:
    # Fecha o navegador
    driver.quit()