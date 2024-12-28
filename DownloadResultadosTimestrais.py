from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Inicia o WebDriver
driver = webdriver.Chrome()
# Maximiza a janela
driver.maximize_window()

# Navega até a página de Relações com Investidores do Magazine Luiza
driver.get('https://ri.magazineluiza.com.br')

# Clica no menu Informações Financeiras
elemento = driver.find_element(By.CSS_SELECTOR, "#heading-mobile-3 > button")
driver.execute_script("arguments[0].click();", elemento)

# Clica em Planilha de Resultado
elemento_planilha = driver.find_element(By.CSS_SELECTOR, "#collapse-mobile-3 > div > ul > li:nth-child(2) > a")
driver.execute_script("arguments[0].click();", elemento_planilha)

# Clica em Planilha de Resultados Trimestrais
elemento_download = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.XPATH, '/html/body/form/main/div[1]/div/div/div/p/a'))
)
driver.execute_script("arguments[0].click();", elemento_download)

# Finalizando o WebDriver
driver.quit()