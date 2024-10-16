from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl
import time

# Configuração do WebDriver com o caminho do ChromeDriver
driver_path = 'C:/Users/maxxi/OneDrive/Área de Trabalho/SK/chromedriver.exe'

# Configurar opções do Chrome
options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)  # Mantém o Chrome aberto após a execução

# Inicializar o WebDriver com as opções
driver = webdriver.Chrome(options=options)

# Acesse o site
driver.get('https://app.mercos.com/391248/modalidades-de-entrega/12707/')

# Aguarde carregar (ajuste o tempo conforme necessário)
time.sleep(2)

# Localize os campos de email e senha e preencha com os dados de login
email_field = driver.find_element(By.XPATH, '//*[@id="id_usuario"]')
email_field.clear()
email_field.send_keys('diretoria@maxximusimportadora.com.br')

senha_field = driver.find_element(By.XPATH, '//*[@id="id_senha"]')
senha_field.clear()
senha_field.send_keys('123456')

# Clique no botão de login
login_button = driver.find_element(By.XPATH, '//*[@id="botaoEfetuarLogin"]')
login_button.click()

# Aguarde a página processar o login (ajuste o tempo conforme necessário)
time.sleep(5)

# Tentar clicar diretamente no botão de regras de modalidade
try:
    button_regras_modalidade = driver.find_element(By.XPATH, '/html/body/div[2]/div[1]/section/div[2]/div/div/div[2]/div[2]/div[1]/div/button')
    driver.execute_script("arguments[0].click();", button_regras_modalidade)
    print("Botão de regras de modalidade clicado com sucesso!")
except Exception as e:
    print(f"Erro ao clicar no botão de regras de modalidade: {e}")

# Aguarde a página processar o clique (ajuste o tempo conforme necessário)
time.sleep(5)

# Caminho do arquivo Excel (dentro da mesma pasta que o script)
excel_file = 'Pasta1.xlsx'

# Abrir o arquivo Excel e ler todas as células de interesse
try:
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook.active
    max_row = sheet.max_row  # Obtém o número total de linhas no arquivo

    for row in range(1, max_row + 1):
        cell_value_A = sheet[f'A{row}'].value
        cell_value_B = sheet[f'B{row}'].value

        # Função para clicar antes de colar o valor
        def clicar_antes_de_colar(xpath_elemento):
            try:
                elemento_clicar = driver.find_element(By.XPATH, xpath_elemento)
                elemento_clicar.click()
                print(f"Clicou antes de colar o valor em {xpath_elemento}.")
            except Exception as e:
                print(f"Erro ao clicar antes de colar o valor em {xpath_elemento}: {e}")

        # Clique e cole o valor da célula A
        clicar_antes_de_colar('//*[@id="modalidade-de-entrega-react-root"]/div/div[2]/div[3]/div[1]/div[2]/div[1]/div[1]/div[2]/div[2]/input')
        try:
            input_field_A = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="modalidade-de-entrega-react-root"]/div/div[2]/div[3]/div[1]/div[2]/div[1]/div[1]/div[2]/div[2]/input'))
            )
            input_field_A.clear()
            input_field_A.send_keys(cell_value_A)
            print(f"Copiado e colado com sucesso na primeira caixa para a linha {row}!")
        except Exception as e:
            print(f"Erro ao preencher o campo na página web para a linha {row}: {e}")

        # Aguarde um tempo para verificar o resultado na página web
        time.sleep(6)

        # Clique e cole o valor da célula B
        clicar_antes_de_colar('//*[@id="modalidade-de-entrega-react-root"]/div/div[2]/div[3]/div[1]/div[2]/div[1]/div[1]/div[2]/div[3]/input')
        try:
            input_field_B = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="modalidade-de-entrega-react-root"]/div/div[2]/div[3]/div[1]/div[2]/div[1]/div[1]/div[2]/div[3]/input'))
            )
            input_field_B.clear()
            input_field_B.send_keys(cell_value_B)
            print(f"Copiado e colado com sucesso na segunda caixa para a linha {row}!")
        except Exception as e:
            print(f"Erro ao preencher o campo na página web para a linha {row}: {e}")

        # Aguarde um tempo para verificar o resultado na página web
        time.sleep(6)

        # Defina os valores nos elementos restantes
        valores = {
            '//*[@id="valorTotalInicio"]': '2000',
            '//*[@id="valorTotalFinal"]': '1000000000',
            '//*[@id="valorFrete"]': '0',
            '//*[@id="prazoEntrega"]': 'CLARO - 8 DIAS'
        }

        for xpath, valor in valores.items():
            try:
                campo = driver.find_element(By.XPATH, xpath)
                campo.clear()
                campo.send_keys(valor)
                print(f"Valor '{valor}' definido com sucesso em {xpath} para a linha {row}!")
            except Exception as e:
                print(f"Erro ao definir o valor '{valor}' em {xpath} para a linha {row}: {e}")

        # Aguarde um tempo para verificar o resultado na página web
        time.sleep(6)

        # Localize e clique no botão 'Confirmar'
        try:
            button_confirmar = driver.find_element(By.XPATH, '/html/body/div[2]/div[1]/section/div[2]/div/div/div[2]/div[3]/div[1]/div[2]/div[2]/button[1]')
            driver.execute_script("arguments[0].click();", button_confirmar)
            print(f"Botão 'Confirmar' clicado com sucesso para a linha {row}!")
        except Exception as e:
            print(f"Erro ao clicar no botão 'Confirmar' para a linha {row}: {e}")

        # Aguarde um tempo para verificar o resultado na página web
        time.sleep(6)

        # Voltar para clicar no botão de regras de modalidade
        try:
            button_regras_modalidade = driver.find_element(By.XPATH, '/html/body/div[2]/div[1]/section/div[2]/div/div/div[2]/div[2]/div[1]/div/button')
            driver.execute_script("arguments[0].click();", button_regras_modalidade)
            print("Botão de regras de modalidade clicado novamente!")
        except Exception as e:
            print(f"Erro ao clicar novamente no botão de regras de modalidade: {e}")

        # Aguarde a página processar o clique (ajuste o tempo conforme necessário)
        time.sleep(5)

except Exception as e:
    print(f"Erro ao abrir o arquivo Excel: {e}")

finally:
    driver.quit()
