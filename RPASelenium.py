#Importação de bibliotecas
import os
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException

# Verificar se o diretório de output já existe
output_dir = './output'
if not os.path.exists(output_dir):
    # Criar o diretório de output
    os.makedirs(output_dir)

# Leitura do arquivo Excel e armazena os dados em um DataFrame usando a biblioteca pandas
df = pd.read_excel('./challenge.xlsx')

# Corrigir espaços extras nos nomes das colunas
df.columns = df.columns.str.strip()

# Configurar o caminho para o Chrome WebDriver
webdriver_service = Service('./chromedriver_win32/chromedriver.exe')  

# Inicializar o driver do Selenium
driver = webdriver.Chrome(service=webdriver_service)

# Abrir o site
driver.get('http://www.rpachallenge.com/')

# Aguardar até que o botão "Start" esteja visível
start_button = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//button[text()="Start"]')))

# Clicar no botão "Start"
start_button.click()

# Aguardar um tempo para que o formulário seja carregado
time.sleep(2)

# A função abaixo é utilizada para localizar elementos na página, permite que o código tente várias vezes antes de desistir e lançar uma exceção.
def find_element_by_xpath_with_retry(xpath):
    MAX_RETRIES = 3
    retries = 0
    while retries < MAX_RETRIES:
        try:
            element = driver.find_element(By.XPATH, xpath)
            return element
        except StaleElementReferenceException:
            retries += 2
            time.sleep(2)
    raise Exception(f"Element not found after {MAX_RETRIES} retries")

# Abrir o arquivo de resultados no diretório de output
resultados_file = os.path.join(output_dir, 'resultados.txt')
with open(resultados_file, 'w') as f:

# Repetir o processo de preenchimento do formulário para cada linha do DataFrame
    for index, row in df.iterrows():
        try:
            # Obter os campos do formulário
            first_name_input = find_element_by_xpath_with_retry('//input[@ng-reflect-name="labelFirstName"]')
            last_name_input = find_element_by_xpath_with_retry('//input[@ng-reflect-name="labelLastName"]')
            company_name_input = find_element_by_xpath_with_retry('//input[@ng-reflect-name="labelCompanyName"]')
            role_input = find_element_by_xpath_with_retry('//input[@ng-reflect-name="labelRole"]')
            address_input = find_element_by_xpath_with_retry('//input[@ng-reflect-name="labelAddress"]')
            email_input = find_element_by_xpath_with_retry('//input[@ng-reflect-name="labelEmail"]')
            phone_input = find_element_by_xpath_with_retry('//input[@ng-reflect-name="labelPhone"]')

            # Preencher os campos do formulário
            first_name_input.clear()
            first_name_input.send_keys(row['First Name'])
            last_name_input.clear()
            last_name_input.send_keys(row['Last Name'])
            company_name_input.clear()
            company_name_input.send_keys(row['Company Name'])
            role_input.clear()
            role_input.send_keys(row['Role in Company'])
            address_input.clear()
            address_input.send_keys(row['Address'])
            email_input.clear()
            email_input.send_keys(row['Email'])
            phone_input.clear()
            phone_input.send_keys(str(row['Phone Number']))

            # Enviar o formulário
            submit_button = find_element_by_xpath_with_retry('//input[@value="Submit"]')
            submit_button.click()

            # Aguardar um tempo para que a página seja atualizada após o envio do formulário
            time.sleep(2)

        except Exception as e:
            print(f"Erro ao preencher formulário para o índice {index}: {str(e)}")

# Capturar uma screenshot da página
screenshot_file = os.path.join(output_dir, f'resultado_{index}.png')
driver.save_screenshot(screenshot_file)

# Capturar o resultado exibido na tela
resultado_element = find_element_by_xpath_with_retry('div[@class="congratulations col s8 m8 18"]')
resultado = resultado_element.text

# Gravar o resultado no arquivo
f.write(f"Resultado para o índice {index}: {resultado}\n")

# Fechar o driver do Selenium
driver.quit()
