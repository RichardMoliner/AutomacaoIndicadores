import os
import locale
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from time import sleep
from datetime import datetime, timedelta
import xlwings as xw

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
sleep(0.5)
brw.find_element(
    By.XPATH, '/html/body/aside/section/div[4]/div/span/ul/li[2]/div/a/span/span[2]').click()
sleep(3)
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

sleep(1)
# Clicar em gerar relatório
brw.find_element(
    By.XPATH, '/html/body/section[1]/div[3]/div[1]/div/div/div/div/form/div[2]/div[2]/button[3]').click()
sleep(5)

# Diretório onde os arquivos estão localizados
diretorio = r"C:\Users\richard.moliner\Downloads"


# Excluir um arquivo específico (substitua 'nome_do_arquivo.txt' pelo nome do arquivo que deseja excluir)
arquivo_a_excluir = os.path.join(diretorio, 'horas_de_ontem.xlsx')
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
    novo_nome = 'horas_de_ontem.xlsx'

    # Renomeie o último arquivo baixado
    os.rename(ultimo_arquivo, os.path.join(diretorio, novo_nome))
    
else:
    print("Nenhum arquivo encontrado para renomear")

# Logout
brw.minimize_window()
brw.find_element(
    By.XPATH, '/html/body/header/div/div[2]/div[2]/div/img').click()
brw.find_element(
    By.XPATH, '/html/body/header/div/div[2]/div[2]/ul/li[12]/a/div').click()
brw.quit()

# Abra o arquivo Excel
path = "C:/Users/richard.moliner/Downloads/horas_de_ontem.xlsx"
wb = xw.Book(path)

# Especifique a planilha e a coluna que você deseja excluir
planilha = wb.sheets['Sheet']  # Altere para o nome da planilha desejada

colunas_a_excluir = ['A:I', 'B:F', 'C:H']  # Colunas a excluir

for coluna_range in colunas_a_excluir:
    coluna = planilha.range(coluna_range)
    coluna.delete()

# Aguarde um momento para as colunas serem excluídas
sleep(2)

# Especifique o intervalo de dados para a tabela dinâmica
intervalo_dados = 'A1:B130'

# Especifique onde deseja que a tabela dinâmica seja criada
celula_tabela_dinamica = 'C2'

# Crie a tabela dinâmica no Excel usando a função Create em um objeto PivotTables
tabela_dinamica_range = planilha.range(intervalo_dados)
pivot_tables = planilha.pivot_tables.add(tabela_dinamica_range, tabledestination=celula_tabela_dinamica)

# Adicione 'Agente' às linhas da tabela dinâmica
pivot_tables.pivotfields('Agente').orientation = xw.constants.PivotFieldOrientation.xlRowField

# Adicione 'Horas Trabalhadas' aos valores da tabela dinâmica (como soma)
campo_horas_trabalhadas = pivot_tables.pivotfields('Horas Trabalhadas')
campo_horas_trabalhadas.orientation = xw.constants.PivotFieldOrientation.xlDataField
campo_horas_trabalhadas.function = xw.constants.PivotFieldFunction.xlSum

# Formate a célula de valores da tabela dinâmica para 'hh:mm:ss'
planilha.range(celula_tabela_dinamica).expand('table').columns[-1].number_format = 'hh:mm:ss'

# Salve o arquivo Excel com a tabela dinâmica
wb.save()

# Feche o arquivo Excel
wb.close()