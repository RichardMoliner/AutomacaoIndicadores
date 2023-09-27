import os
import locale
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from time import sleep
from datetime import datetime, timedelta
from selenium.webdriver.common.action_chains import ActionChains

# Setando timezone para testo do dia da semana
locale.setlocale(locale.LC_TIME, 'pt_BR.utf8')

# Lista dos agentes
agentes = ['Richard Moliner Júnior',
           'Cristhielle Pessoa',
           'Micheli Machado',
           'Vitoria Longaretti',
           'Douglas Gonçalves e Barra',
           'Pedro Henrique Aurélio Martins',
           'Juliana Scorsatto',
           'Keity Santos',
           'Tálles Passaura de Mattos']



# Constants
ACCOUNT = ''
PASSWORD = ''
URL = 'https://atendimento.sistemainfo.com.br/Account/Login'


# Login
brw = webdriver.Chrome()
brw.get(URL)
brw.find_element(
    By.XPATH, '/html/body/div[1]/section/div/div[1]/div/div[2]/div[2]/div[2]/form/div/div/div[1]/input[1]').send_keys(ACCOUNT)
brw.find_element(
    By.XPATH, '/html/body/div[1]/section/div/div[1]/div/div[2]/div[2]/div[2]/form/div/div/div[2]/input[1]').send_keys(PASSWORD)
brw.find_element(
    By.XPATH, '/html/body/div[1]/section/div/div[1]/div/div[2]/div[2]/div[2]/form/div/div/div[4]/button').click()
sleep(3)
brw.maximize_window()

# Abertura da página do relatório
brw.find_element(
    By.XPATH, '/html/body/aside/section/div[4]/div/span/div/span[1]/i').click()
brw.find_element(
    By.XPATH, '/html/body/aside/section/div[4]/div/span/ul/li[2]/div/a/span/span[2]').click()
sleep(2)
brw.find_element(
    By.XPATH, '/html/body/section[1]/div[3]/div[1]/div/div/div/div/div[2]/a[2]/div/div/h6').click()
sleep(2)

# Função para validação se o dia anterior é útil e armazenar
def obter_dia_de_ontem():
    hoje = datetime.now()

    dia_de_ontem = hoje - timedelta(days=1)


    while dia_de_ontem.weekday() >= 5:
        dia_de_ontem -= timedelta(days=1)
    dia_da_semana = dia_de_ontem.strftime('%A')
    dia_de_ontem = dia_de_ontem.strftime('%d')

    return dia_de_ontem, dia_da_semana


dia_anterior = obter_dia_de_ontem()[0]
dia_anterior_texto = obter_dia_de_ontem()[1]

# Setar a data de início
data_inicio = brw.find_element(
    By.XPATH, '/html/body/section[1]/div[3]/div[1]/div/div/div/div/form/div[1]/div/div[2]/div/div/div/div[1]/div[1]/div/input').click()
dias_inicio = brw.find_elements(
    By.XPATH, "//p[@class='reports-MuiTypography-root reports-MuiTypography-body2 reports-MuiTypography-colorInherit']")

for dia in dias_inicio:
    valor_dia = dia.text
    if valor_dia == dia_anterior:
        dia.click()
        break
sleep(1)

# Setar a data de fim
data_fim = brw.find_element(
    By.XPATH, '/html/body/section[1]/div[3]/div[1]/div/div/div/div/form/div[1]/div/div[2]/div/div/div/div[1]/div[2]/div/input').click()
dias_fim = brw.find_elements(
    By.XPATH, "//p[@class='reports-MuiTypography-root reports-MuiTypography-body2 reports-MuiTypography-colorInherit']")

for dia in dias_fim:
    valor_dia = dia.text
    if valor_dia == dia_anterior:
        dia.click()
        break

# Buscar todos os agentes da lista
lista_agentes = brw.find_element(
    By.XPATH, '/html/body/section[1]/div[3]/div[1]/div/div/div/div/form/div[1]/div/div[2]/div/div/div/div[3]/div[1]/div/div/div/input')
lista_agentes.click()

for agente in agentes:
    lista_agentes.send_keys(agente)
    sleep(1)
    lista_agentes.send_keys(Keys.ENTER)


# Clicar em gerar relatório
brw.find_element(
    By.XPATH, '/html/body/section[1]/div[3]/div[1]/div/div/div/div/form/div[2]/div[2]/button[3]').click()
sleep(5)

# Diretório onde os arquivos estão localizados
diretorio = r"C:\Users\richard.moliner\Downloads"


# Excluir um arquivo específico (substitua 'nome_do_arquivo.txt' pelo nome do arquivo que deseja excluir)
arquivo_a_excluir = os.path.join(diretorio, 'horas de ontem.xlsx')
if os.path.exists(arquivo_a_excluir):
    os.remove(arquivo_a_excluir)
    print(f"O arquivo {arquivo_a_excluir} foi excluído.")
else:
    print(f"O arquivo {arquivo_a_excluir} não foi encontrado.")

# Exportar e selecionar XSLS
brw.find_element(By.XPATH, "//span[text()='Exportar']").click()
sleep(0.5)
brw.find_element(By.XPATH, '/html/body/div[8]/div[3]/div/form/div[2]/div[2]/label[2]/span[2]/div/p').click()
brw.find_element(By.XPATH, '/html/body/div[8]/div[3]/div/form/div[3]/div/button[2]/span[1]').click()
sleep(5)


# Lista todos os arquivos no diretório
arquivos = os.listdir(diretorio)

# Verifique o último arquivo baixado com base na data de modificação
ultimo_arquivo = None
data_modificacao = datetime(1970, 1, 1)
for arquivo in arquivos:
    caminho_completo = os.path.join(diretorio, arquivo)
    data = datetime.fromtimestamp(os.path.getmtime(caminho_completo))
    if data > data_modificacao:
        data_modificacao = data
        ultimo_arquivo = caminho_completo

# Verifique se encontramos um arquivo para renomear
if ultimo_arquivo:
    # Crie uma data para 'horas de ontem'
    novo_nome = 'horas de ontem.xlsx'

    # Renomeie o último arquivo baixado
    os.rename(ultimo_arquivo, os.path.join(diretorio, novo_nome))
    
else:
    print("Nenhum arquivo encontrado para renomear")

# Logout
brw.find_element(
    By.XPATH, '/html/body/header/div/div[2]/div[2]/div/img').click()
brw.find_element(
    By.XPATH, '/html/body/header/div/div[2]/div[2]/ul/li[12]/a/div').click()
sleep(1)
brw.quit()
