from datetime import datetime
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException
import secrets
import string
import time

chrome_options = Options()
chrome_options.add_argument("--incognito")
driver = webdriver.Chrome(options=chrome_options)

status, emp_num, link, dados, titulos, emails, empresas = []

date = datetime.now().strftime("%d-%m-%Y")
df = pd.read_excel('email.xlsx')

#Função responsável por criar automáticamente senhas reforçadas para o login
def generate_random_password():
    characters = string.ascii_letters + string.digits + "!@#$%&*"
    password_length = secrets.choice([11, 12, 13])
    password = ''.join(secrets.choice(characters) for i in range(password_length))
    return password

#Se autenticar no site com login e senha
try:
    driver.get("https://www.********.com.br/gestaoSite/empresas")

    campo_usuario = WebDriverWait(driver, 3).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'div.imagemDir form ul li input.codigo.validar'))
    )
    campo_senha = WebDriverWait(driver, 3).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'div.imagemDir form ul li input.senha'))
    )

    campo_usuario.send_keys('login')
    campo_senha.send_keys('*****')

    botao_login = driver.find_element(By.CSS_SELECTOR, 'div.imagemDir form ul li a.bt')  
    botao_login.click()

except NoSuchElementException as e:
    print("Erro ao localizar os elementos da página:", e)
    driver.quit()
    exit()


#Criando as variaveis com as informações referentes para cada novo cadastro
for index, row in df.iterrows():
    empresatext = row['Empresa']
    titulotext = row['Link1_Titulo']
    slugtext = row['Link1_Slug']
    emailtext = row['EMAIL']
    codigotext = row['FORMULARIO']
    senhatext = generate_random_password()


    #Acessando o link de administração de usuarios
    link = "https://www.*******.com.br/gestaoSite/empresas/edit/"
    driver.get(link)

    #Acessando cada item para futuro preenchimento
    nome = WebDriverWait(driver, 3).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'input[name="nome"]'))
    )

    email = WebDriverWait(driver, 3).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'input[name="email"]'))
    )

    status = WebDriverWait(driver, 3).until(
        EC.presence_of_element_located((By.NAME, 'status'))
    )
    
    #Selecionando item no combo box
    stats = Select(status)
    stats.select_by_value('1')


    titulo = WebDriverWait(driver, 3).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'input[name="tituloLink1"]'))
    )

    slug = WebDriverWait(driver, 3).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'input[name="slugLink1"]'))
    )

    url = WebDriverWait(driver, 3).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'textarea[name="urlLink1"]'))
    )

    tipolink = WebDriverWait(driver, 3).until(
        EC.presence_of_element_located((By.NAME, 'tipoLink1'))
    )

    link = Select(tipolink)
    link.select_by_value('url')

    senha = WebDriverWait(driver, 3).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'input[name="senha"]'))
    )


    senha2 = WebDriverWait(driver, 3).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'input[name="csenha"]'))
    )


    #Enviando as variaveis salvas anteriormente para os inputs (campos de texto) p/ preenchimento obrigatório 
    nome.send_keys(empresatext)
    email.send_keys(emailtext)
    titulo.send_keys(titulotext)
    slug.send_keys(slugtext)
    url.send_keys(codigotext)
    senha.send_keys(senhatext)
    senha2.send_keys(senhatext)


    #Clicando no botão de salvar e dando um sleep evitar sobrecarregar
    time.sleep(5)
    botao_login = driver.find_element(By.CSS_SELECTOR, 'a[class="bt"]')  
    botao_login.click()
 

driver.quit()
