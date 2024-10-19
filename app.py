# selenium 4
import openpyxl
import pyperclip
from time import sleep
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

def iniciar_driver():
    chrome_options = Options()
    # Fonte de opções de switches https://peter.sh/experiments/chromium-command-line-switches/

    arguments = ['--lang=pt-BR', '--window-size=800,600',
                '--incognito']
    ''' Common arguments
    --start-maximized # Inicia maximizado
    --lang=pt-BR # Define o idioma de inicialização, # en-us , pt-BR
    --incognito # Usar o modo anônimo
    --window-size=800,800 # Define a resolução da janela em largura e altura
    --headless # Roda em segundo plano(com a janela fechada)
    --disable-notifications # Desabilita notificações
    --disable-gpu # Desabilita renderização com GPU
    '''
    for argument in arguments:
        chrome_options.add_argument(argument)

    caminho_padrao_para_download = 'E:\\Storage\\Desktop'

    # Lista de opções experimentais(nem todas estão documentadas) https://chromium.googlesource.com/chromium/src/+/32352ad08ee673a4d43e8593ce988b224f6482d3/chrome/common/pref_names.cc
    chrome_options.add_experimental_option("prefs", {
        'download.default_directory': caminho_padrao_para_download,
        # Atualiza diretório para diretório passado acima
        'download.directory_upgrade': True,
        # Setar se o navegar deve pedir ou não para fazer download
        'download.prompt_for_download': False,
        "profile.default_content_setting_values.notifications": 2,  # Desabilita notificações
        # Permite realizar múltiplos downlaods multiple downloads
        "profile.default_content_setting_values.automatic_downloads": 1,
    })

    driver = webdriver.Chrome(options=chrome_options)
    return driver

driver = iniciar_driver()
driver.get('https://carlosfortunatodev.github.io/Cadastro-Fortuna-Tech/')
sleep(5)

workbook = openpyxl.load_workbook('funcionarios.xlsx')
sheet_funcionario = workbook['Sheet1']
for linha in sheet_funcionario.iter_rows(min_row=2):

# Nome Completo
    nome_completo = linha[0].value
    pyperclip.copy(nome_completo)
    # encontrar e clicar no campo de nome
    campo_nome = driver.find_element(By.ID, 'fullName')
    sleep(1)
    # digitar o nome
    campo_nome.send_keys(nome_completo)
    sleep(1)

# Chapa
    chapa = linha[1].value
    pyperclip.copy(chapa)
    # encontrar e clicar no campo chapa
    campo_chapa = driver.find_element(By.ID, 'chapa')
    sleep(1)
    # digitar o nome
    campo_chapa.send_keys(chapa)
    sleep(1)

# Rg
    rg = linha[2].value
    pyperclip.copy(rg)
    # encontrar e clicar no campo rg
    campo_rg = driver.find_element(By.ID, 'rg')
    sleep(1)
    # digitar o nome
    campo_rg.send_keys(rg)
    sleep(1)

# Salario Base
    salarioBase = linha[3].value
    pyperclip.copy(salarioBase)
    # encontrar e clicar no campo rg
    campo_salario_base = driver.find_element(By.ID, 'salarioBase')
    sleep(1)
    # digitar o nome
    campo_salario_base.send_keys(salarioBase)
    sleep(1)

    # Rolar tela para baixo
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    sleep(1)

# Jornada De Trabalho
    jornadaDeTrabalho = linha[4].value
    pyperclip.copy(jornadaDeTrabalho)
    # encontrar e clicar no campo rg
    campo_jornada_de_trabalho = driver.find_element(By.ID, 'jornadaDeTrabalho')
    sleep(1)
    # digitar o nome
    campo_jornada_de_trabalho.send_keys(jornadaDeTrabalho)
    sleep(1)

# Horas Trabalhadas
    horasTrabalhadas = linha[5].value
    pyperclip.copy(horasTrabalhadas)
    # encontrar e clicar no campo rg
    campo_horas_trabalhadas = driver.find_element(By.ID, 'horasTrabalhadas')
    sleep(1)
    # digitar o nome
    campo_horas_trabalhadas.send_keys(horasTrabalhadas)
    sleep(1)

# Hora Extra
    horaExtra = linha[6].value
    pyperclip.copy(horaExtra)
    # encontrar e clicar no campo rg
    campo_hora_extra = driver.find_element(By.ID, 'horaExtra')
    sleep(1)
    # digitar o nome
    campo_hora_extra.send_keys(horaExtra)
    sleep(1)

    # De acordo (check)
    check = driver.find_element(By.ID, 'check')
    check.click()
    sleep(1)

    # Enviar
    button_enviar = driver.find_element(By.XPATH, "//button[text()='Enviar']")
    button_enviar.click()
    sleep(1)

    # Ok
    button_ok = driver.find_element(By.ID, 'reset')
    button_ok.click()
    sleep(1)

    # Rolar tela para baixo
    driver.execute_script("window.scrollTo(0, document.body.scrollTop);")
    sleep(1)


input('')