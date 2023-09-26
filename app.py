from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from time import sleep
from datetime import datetime, timedelta
import locale
from selenium.webdriver.common.action_chains import ActionChains

agentes = ['Richard']

locale.setlocale(locale.LC_TIME, 'pt_BR.utf8')

ACCOUNT = 'richard'
PASSWORD = 'Ririchard.1'
URL = 'https://atendimento.sistemainfo.com.br/Account/Login'


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
brw.find_element(
    By.XPATH, '/html/body/aside/section/div[4]/div/span/div/span[1]/i').click()
brw.find_element(
    By.XPATH, '/html/body/aside/section/div[4]/div/span/ul/li[2]/div/a/span/span[2]').click()
sleep(2)
brw.find_element(
    By.XPATH, '/html/body/section[1]/div[3]/div[1]/div/div/div/div/div[2]/a[2]/div/div/h6').click()
sleep(2)


def obter_dia_de_ontem():
    hoje = datetime.now()

    dia_de_ontem = hoje - timedelta(days=1)

    # Verifica se o dia de ontem é sábado (5) ou domingo (6)
    while dia_de_ontem.weekday() >= 5:
        dia_de_ontem -= timedelta(days=1)
    dia_da_semana = dia_de_ontem.strftime('%A')
    dia_de_ontem = dia_de_ontem.strftime('%d')

    return dia_de_ontem, dia_da_semana


dia_anterior = obter_dia_de_ontem()[0]
dia_anterior_texto = obter_dia_de_ontem()[1]

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

data_fim = brw.find_element(
    By.XPATH, '/html/body/section[1]/div[3]/div[1]/div/div/div/div/form/div[1]/div/div[2]/div/div/div/div[1]/div[2]/div/input').click()
dias_fim = brw.find_elements(
    By.XPATH, "//p[@class='reports-MuiTypography-root reports-MuiTypography-body2 reports-MuiTypography-colorInherit']")

for dia in dias_fim:
    valor_dia = dia.text
    if valor_dia == dia_anterior:
        dia.click()
        break

lista_agentes = brw.find_element(
    By.XPATH, '/html/body/section[1]/div[3]/div[1]/div/div/div/div/form/div[1]/div/div[2]/div/div/div/div[3]/div[1]/div/div/div/input')
lista_agentes.click()

for agente in agentes:
    lista_agentes.send_keys(agente)
    sleep(1)
    lista_agentes.send_keys(Keys.ENTER)


brw.find_element(
    By.XPATH, '/html/body/section[1]/div[3]/div[1]/div/div/div/div/form/div[2]/div[2]/button[3]').click()
sleep(3)

src = brw.find_element(
    By.XPATH, "//div[@class='ReactVirtualized__Grid']//div[@class='ReactVirtualized__Grid__innerScrollContainer']//div[@data-index='9']")
sleep(1)
dest = brw.find_element(
    By.XPATH, "//div[@data-testid='drop-zone']//p[@class='reports-MuiTypography-root sc-dkrFOg bpDwEs sc-iOeugr bpTxYL reports-MuiTypography-body1']")
sleep(1)

action_chains = ActionChains(brw)
sleep(1)

action_chains.drag_and_drop(src, dest)

brw.find_element(
    By.XPATH, '/html/body/header/div/div[2]/div[2]/div/img').click()
brw.find_element(
    By.XPATH, '/html/body/header/div/div[2]/div[2]/ul/li[12]/a/div').click()
sleep(1)
brw.quit()
