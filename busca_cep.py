from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook  
import subprocess
import pyautogui as py

py.PAUSE = 0.5
FAILSAFE = True

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)
wait = WebDriverWait(driver, 20)

driver.get('https://buscacepinter.correios.com.br/app/endereco/index.php?t')

#Pesquisar CEP e apertar em "Buscar"
wait.until(EC.visibility_of_element_located(('name', 'endereco'))).send_keys('88340079')
wait.until(EC.visibility_of_element_located(('name', 'btn_pesquisar'))).click()

#Pegar o caminho + o nome do arquivo do computador
planilha = '/home/will/Documentos/GitHub/busca_cep/Pesquisa endereços.xlsx'
planilha_dados = load_workbook(planilha)

#Selecionar a aba CEP de dentro da minha planilha
sheet_cep = planilha_dados['CEP']

#Pegar a última linha preenchida e acrescentar +1
linha_cep = len(sheet_cep['A']) + 1

for i in range(2, linha_cep):
    
    #Clicar em Nova busca
    wait.until(EC.visibility_of_element_located(('xpath', '//*[@id="btn_nbusca"]'))).click()
    
    #Pegar o CEP da planilha de jogar no site
    cep_pesquisa = sheet_cep['A%s' % i].value
    # %s = Transforma para string
    # % = Operador modulo. Ele retorna o resto da divisão do operando da esquerda pelo operando da direita.
    # .value = Para pegar o valor da celula
    
    #Pesquisar CEP e apertar em "Buscar"
    wait.until(EC.visibility_of_element_located(('name', 'endereco'))).send_keys(cep_pesquisa)
    wait.until(EC.visibility_of_element_located(('name', 'btn_pesquisar'))).click()
    
    #Extrair dados pesquisados
    rua = wait.until(EC.visibility_of_element_located(('xpath', '//*[@id="resultado-DNEC"]/tbody/tr/td[1]'))).text
    bairro = wait.until(EC.visibility_of_element_located(('xpath', '//*[@id="resultado-DNEC"]/tbody/tr/td[2]'))).text
    cidade = wait.until(EC.visibility_of_element_located(('xpath', '//*[@id="resultado-DNEC"]/tbody/tr/td[3]'))).text
    cep = wait.until(EC.visibility_of_element_located(('xpath', '//*[@id="resultado-DNEC"]/tbody/tr/td[4]'))).text

    print(rua, bairro, cidade, cep, sep=', ')
    
    #Seleciona a aba Dados dentro da minha planilha
    sheet_dados = planilha_dados['Dados']
    
    #Pegar a última linha preenchida e acrescentar +1
    linha_dados = len(sheet_dados['A']) + 1
    
    #Criar a variável para juntar A + a última linha. Ex: A2
    colunaA = 'A' + str(linha_dados) 
    colunaB = 'B' + str(linha_dados)
    colunaC = 'C' + str(linha_dados)
    colunaD = 'D' + str(linha_dados)

    #Colando as informações extraídas do site na planilha
    sheet_dados[colunaA] = rua #A2
    sheet_dados[colunaB] = bairro #B2
    sheet_dados[colunaC] = cidade #C2
    sheet_dados[colunaD] = cep #D2 

#Salvar planilha com as novas informações
planilha_dados.save(filename=planilha)

#Abrir a planilha
subprocess.run(['xdg-open', planilha])

py.alert('O BOT foi executado com exito!')
driver.quit()

