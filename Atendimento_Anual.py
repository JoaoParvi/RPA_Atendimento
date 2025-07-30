import time
import pandas as pd
import urllib 
from selenium.webdriver.chrome.options import Options
from datetime import datetime
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from sqlalchemy import create_engine
from datetime import date
from selenium.webdriver.common.keys import Keys

# Configurar o WebDriver usando o webdriver_manager para sempre baixar a versão mais atualizada
navegador = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# Definir a URL e as credenciais
url = 'https://www.app-indecx.com/'
login = "ericka.nascimento@parvi.com.br"
senha = "Ericka@123"

# Navegar até a URL
navegador.get(url)
navegador.maximize_window()

# Função auxiliar para enviar múltiplas teclas
def send_multiple_keys(navegador, key, times):
    for _ in range(times):
        navegador.switch_to.active_element.send_keys(key)
        time.sleep(1)

options = Options()
options.add_argument("start-maximized")  # opcional
options.add_argument("window-size=1920,1080") 
# LOGIN
campo_login = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#email")))
campo_login.click()
campo_login.send_keys(login)

campo_senha = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#password")))
campo_senha.click()
campo_senha.send_keys(senha)

botaologin = WebDriverWait(navegador, 20).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".btn-block > span:nth-child(1)")))
botaologin.click()
time.sleep(5)

navegador.execute_script("document.body.style.zoom='90%'")

time.sleep(10)
time.sleep(4)

# ACESSAR RELATÓRIOS
relatorios = WebDriverWait(navegador, 20).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, "#app > section > section > aside > div > ul > li:nth-child(6) > div > span > i"))
)
relatorios.click()
time.sleep(4)

listarelatorio = WebDriverWait(navegador, 20).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, "#sub_menu_4_\$\$_menu-sub-item-dashboard\.sideMenu\.reports\.title-popup > li:nth-child(1) > span > span"))
)
listarelatorio.click()
time.sleep(4)

CliqueOutros = WebDriverWait(navegador, 20).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, "#pdf-dashboard > div > div > ul > li:nth-child(8) > div > span > span"))
)
CliqueOutros.click()
time.sleep(4)

navegador.switch_to.active_element.send_keys(Keys.PAGE_DOWN)
time.sleep(4)
navegador.switch_to.active_element.send_keys(Keys.PAGE_DOWN)
time.sleep(4)
navegador.switch_to.active_element.send_keys(Keys.PAGE_DOWN)
time.sleep(10)


relatorioCXI = WebDriverWait(navegador, 20).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, "#sub_menu_24_\$\$_menu-sub-item-reportSession\.reportList\.menu\.others\.title-popup > li:nth-child(6) > span > span"))
)
relatorioCXI.click()
time.sleep(4)

# APLICAR FILTROS (VENDAS DE CAMINHÃO)
filtroData = WebDriverWait(navegador, 20).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, "#header-filter-button"))
)
filtroData.click()
time.sleep(4)

filtroSegmento = WebDriverWait(navegador, 20).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter-container > div:nth-child(3) > div.ant-select.default-select.--block.--medium.filter__action-group-select.ant-select-single.ant-select-allow-clear.ant-select-show-arrow.ant-select-show-search > div > span.ant-select-selection-item"))
)
filtroSegmento.click()
time.sleep(4)

navegador.switch_to.active_element.send_keys(Keys.ARROW_DOWN)
time.sleep(4)
navegador.switch_to.active_element.send_keys(Keys.ARROW_DOWN)
time.sleep(4)
navegador.switch_to.active_element.send_keys(Keys.ARROW_DOWN)
time.sleep(4)
navegador.switch_to.active_element.send_keys(Keys.ENTER)
time.sleep(4)

###########################################################################################################
# selecionando loja
navegador.switch_to.active_element.send_keys(Keys.TAB)
time.sleep(4)
navegador.switch_to.active_element.send_keys(Keys.TAB)
time.sleep(4)
# Envia a letra "m" para o campo atualmente em foco
navegador.switch_to.active_element.send_keys("m")
time.sleep(4)
navegador.switch_to.active_element.send_keys(Keys.ENTER)
time.sleep(4)
navegador.switch_to.active_element.send_keys(Keys.TAB)
time.sleep(4)
############################################################################################################
# Aplicar o filtro final
Aplicar = WebDriverWait(navegador, 20).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter__apply-btn"))
)
Aplicar.click()
time.sleep(2)
navegador.switch_to.active_element.send_keys(Keys.PAGE_UP)
time.sleep(4)
navegador.switch_to.active_element.send_keys(Keys.PAGE_UP)
time.sleep(4)
navegador.switch_to.active_element.send_keys(Keys.PAGE_UP)
time.sleep(4)

print("Page Up realizado com sucesso.")
################################################### FILTRO DATA #########################################################################################
CliqueData = WebDriverWait(navegador, 20).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, "#pdf-dashboard > div > main > section.default-card-container.--undefined.--padding-medium.--justify-between.--align-undefined.--direction-undefined.--bg-type-primary > div > div > div"))
)

print ("Clique Input de Data")

CliqueData.click()
time.sleep(4)
#clique principal
####################
CliqueAtual = WebDriverWait(navegador, 20).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, "body > div:nth-child(7) > div > div > div > div.ant-picker-panel-container > div.ant-picker-footer > ul > li:nth-child(8) > span"))
)

print ("Primeiro Clique")

CliqueAtual.click()
time.sleep(5)
###################
Aplicardata = WebDriverWait(navegador, 20).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, "body > div:nth-child(7) > div > div > div > div.ant-picker-panel-container > div.ant-picker-footer > div > div > button.ant-btn.ant-btn-primary.default-btn.--font-default.--primary.--undefined.--medium"))
)
Aplicardata.click()
time.sleep(5)

print ("Clique de Confirmação")

# Aplicardata = WebDriverWait(navegador, 20).until(
#     EC.element_to_be_clickable((By.CSS_SELECTOR, "body > div:nth-child(7) > div > div > div > div.ant-picker-panel-container > div.ant-picker-footer > div > div > button.ant-btn.ant-btn-primary.default-btn.--font-default.--primary.--undefined.--medium"))
# )
# Aplicardata.click()
# time.sleep(2)

# print ("Aplicando")
#######################################################################################################################################################


# coleta de dados boa vista

# Coletar dados do relatório
NotaNacionalAtn = WebDriverWait(navegador, 20).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, "#pdf-dashboard > div > main > section:nth-child(3) > div > span:nth-child(1) > strong"))
)
Nota_NacionalAtn = NotaNacionalAtn.text
time.sleep(2)

NotafilialBoaVista = WebDriverWait(navegador, 20).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, "#pdf-dashboard > div > main > div:nth-child(4) > div.ant-col.ant-col-8 > section > div > div > div > div > span"))
)
NotafilialBoaVista = NotafilialBoaVista.text
time.sleep(2)

NotaNacionalEquipeVendas = WebDriverWait(navegador, 20).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, "#pdf-dashboard > div > main > div:nth-child(8) > div.ant-col.ant-col-8 > section > div > div > div > div > span"))
)
Nota_Nacional_Equipe_Vendas = NotaNacionalEquipeVendas.text
time.sleep(2)

NotafilialEqVendasBoavista = WebDriverWait(navegador, 20).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, "#pdf-dashboard > div > main > div:nth-child(8) > div.ant-col.ant-col-8 > section > div > div > div > div > span"))
)
Nota_filial_EqVendas_Boavista = NotafilialEqVendasBoavista.text
time.sleep(2)

# Exibir os dados
print(f"ATENDIMENTO AO CLIENTE BOA VISTA: {NotafilialBoaVista}")
print(f"NOTA NACIONAL ATENDIMENTO AO CLIENTE: {Nota_NacionalAtn}")
print(f"NOTA NACIONAL EQUIPE VENDAS: {Nota_Nacional_Equipe_Vendas}")
print(f"NOTA FILIAL EQVENDAS BOAVISTA: {Nota_filial_EqVendas_Boavista}")
########################################################################

# Clicar no botão de filtro de data
filtroData = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#header-filter-button"))
)
filtroData.click()
time.sleep(4)


# Selecionar "Boavista"
desmarcar = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter-container > div:nth-child(4) > div.ant-select.default-select.--block.--medium.filter__branch-select.ant-select-multiple.ant-select-allow-clear.ant-select-show-search > div > div > div:nth-child(1) > span > span.ant-select-selection-item-remove > span > svg > path"))
)
desmarcar.click()
time.sleep(4)

  # Selecionar "Brasilia"
filtroConcessionaria = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter-container > div:nth-child(4) > div.ant-select.default-select.--block.--medium.filter__branch-select.ant-select-multiple.ant-select-allow-clear.ant-select-show-search > div > div"))
)
filtroConcessionaria.click()
time.sleep(4)

send_multiple_keys(navegador, Keys.ARROW_DOWN, 1)
time.sleep(2)

send_multiple_keys(navegador, Keys.ENTER, 1)
time.sleep(2)

# Confirmar "Brasilia"
AplicarFiltro = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter__apply-btn"))
)
AplicarFiltro.click()
time.sleep(4)

# Coletar dados do relatório

NotafilialBrasilia = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#pdf-dashboard > div > main > div:nth-child(4) > div.ant-col.ant-col-8 > section > div > div > div > div > span")))
NotafilialBrasilia = NotafilialBrasilia.text  # Extrair o texto do elemento
time.sleep(2)

NotafilialEqVendasBrasilia = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#pdf-dashboard > div > main > div:nth-child(8) > div.ant-col.ant-col-8 > section > div > div > div > div > span")))
Nota_filial_EqVendas_Brasilia = NotafilialEqVendasBrasilia.text  # Extrair o texto do elemento
time.sleep(2)


print(f"ATENDIMENTO AO CLIENTE BRASILIA: {NotafilialBrasilia}")
print(f"NOTA FILIAL EQVENDAS BRASILIA: {Nota_filial_EqVendas_Brasilia}")
#######################################

# Clicar no botão de filtro de data
filtroData = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#header-filter-button"))
)
filtroData.click()
time.sleep(4)


# Selecionar "Boavista"
desmarcar = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter-container > div:nth-child(4) > div.ant-select.default-select.--block.--medium.filter__branch-select.ant-select-multiple.ant-select-allow-clear.ant-select-show-search > div > div > div:nth-child(1) > span > span.ant-select-selection-item-remove > span > svg > path"))
)
desmarcar.click()
time.sleep(4)

  # Selecionar "Boavista"
filtroConcessionaria = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter-container > div:nth-child(4) > div.ant-select.default-select.--block.--medium.filter__branch-select.ant-select-multiple.ant-select-allow-clear.ant-select-show-search > div > div"))
)
filtroConcessionaria.click()
time.sleep(4)

send_multiple_keys(navegador, Keys.ARROW_DOWN, 2)
time.sleep(2)

send_multiple_keys(navegador, Keys.ENTER, 1)
time.sleep(2)

# Confirmar "Boavista"
AplicarFiltro = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter__apply-btn"))
)
AplicarFiltro.click()
time.sleep(4)

NotafilialCampos = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#pdf-dashboard > div > main > div:nth-child(4) > div.ant-col.ant-col-8 > section > div > div > div > div > span")))
NotafilialCampos = NotafilialCampos.text  # Extrair o texto do elemento
time.sleep(2)

NotafilialEqVendasCampos = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#pdf-dashboard > div > main > div:nth-child(8) > div.ant-col.ant-col-8 > section > div > div > div > div > span")))
Nota_filial_EqVendas_Campos = NotafilialEqVendasCampos.text  # Extrair o texto do elemento
time.sleep(2)

print(f"ATENDIMENTO AO CLIENTE CAMPOS: {NotafilialCampos}")
print(f"NOTA FILIAL EQVENDAS CAMPOS: {Nota_filial_EqVendas_Campos}")

# Clicar no botão de filtro de data
filtroData = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#header-filter-button"))
)
filtroData.click()
time.sleep(4)


# Selecionar "Boavista"
desmarcar = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter-container > div:nth-child(4) > div.ant-select.default-select.--block.--medium.filter__branch-select.ant-select-multiple.ant-select-allow-clear.ant-select-show-search > div > div > div:nth-child(1) > span > span.ant-select-selection-item-remove > span > svg > path"))
)
desmarcar.click()
time.sleep(4)

  # Selecionar "Boavista"
filtroConcessionaria = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter-container > div:nth-child(4) > div.ant-select.default-select.--block.--medium.filter__branch-select.ant-select-multiple.ant-select-allow-clear.ant-select-show-search > div > div"))
)
filtroConcessionaria.click()
time.sleep(4)

send_multiple_keys(navegador, Keys.ARROW_DOWN, 3)
time.sleep(2)

send_multiple_keys(navegador, Keys.ENTER, 1)
time.sleep(2)

# Confirmar "Boavista"
AplicarFiltro = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter__apply-btn"))
)
AplicarFiltro.click()
time.sleep(4)

NotafilialFloriano = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#pdf-dashboard > div > main > div:nth-child(4) > div.ant-col.ant-col-8 > section > div > div > div > div > span")))
NotafilialFloriano = NotafilialFloriano.text  # Extrair o texto do elemento
time.sleep(2)

NotafilialEqVendasFloriano = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#pdf-dashboard > div > main > div:nth-child(8) > div.ant-col.ant-col-8 > section > div > div > div > div > span")))
Nota_filial_EqVendas_Floriano = NotafilialEqVendasFloriano.text  # Extrair o texto do elemento
time.sleep(2)

print(f"ATENDIMENTO AO CLIENTE FLORIANO: {NotafilialFloriano}")
print(f"NOTA FILIAL EQVENDAS FLORIANO: {Nota_filial_EqVendas_Floriano}")

# Clicar no botão de filtro de data
filtroData = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#header-filter-button"))
)
filtroData.click()
time.sleep(4)


# Selecionar "Boavista"
desmarcar = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter-container > div:nth-child(4) > div.ant-select.default-select.--block.--medium.filter__branch-select.ant-select-multiple.ant-select-allow-clear.ant-select-show-search > div > div > div:nth-child(1) > span > span.ant-select-selection-item-remove > span > svg > path"))
)
desmarcar.click()
time.sleep(4)

  # Selecionar "Boavista"
filtroConcessionaria = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter-container > div:nth-child(4) > div.ant-select.default-select.--block.--medium.filter__branch-select.ant-select-multiple.ant-select-allow-clear.ant-select-show-search > div > div"))
)
filtroConcessionaria.click()
time.sleep(4)

send_multiple_keys(navegador, Keys.ARROW_DOWN, 4)
time.sleep(2)

send_multiple_keys(navegador, Keys.ENTER, 1)
time.sleep(2)

# Confirmar "Boavista"
AplicarFiltro = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter__apply-btn"))
)
AplicarFiltro.click()
time.sleep(4)

NotafilialLuziana = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#pdf-dashboard > div > main > div:nth-child(4) > div.ant-col.ant-col-8 > section > div > div > div > div > span")))
NotafilialLuziana = NotafilialLuziana.text  # Extrair o texto do elemento
time.sleep(2)

NotafilialEqVendasLuziana = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#pdf-dashboard > div > main > div:nth-child(8) > div.ant-col.ant-col-8 > section > div > div > div > div > span")))
Nota_filial_EqVendas_Luziana = NotafilialEqVendasLuziana.text  # Extrair o texto do elemento
time.sleep(2)

print(f"ATENDIMENTO AO CLIENTE LUZIANA: {NotafilialLuziana}")
print(f"NOTA FILIAL EQVENDAS LUZIANA: {Nota_filial_EqVendas_Luziana}")

# Clicar no botão de filtro de data
filtroData = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#header-filter-button"))
)
filtroData.click()
time.sleep(4)


# Selecionar "Boavista"
desmarcar = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter-container > div:nth-child(4) > div.ant-select.default-select.--block.--medium.filter__branch-select.ant-select-multiple.ant-select-allow-clear.ant-select-show-search > div > div > div:nth-child(1) > span > span.ant-select-selection-item-remove > span > svg > path"))
)
desmarcar.click()
time.sleep(4)

  # Selecionar "Boavista"
filtroConcessionaria = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter-container > div:nth-child(4) > div.ant-select.default-select.--block.--medium.filter__branch-select.ant-select-multiple.ant-select-allow-clear.ant-select-show-search > div > div"))
)
filtroConcessionaria.click()
time.sleep(4)

send_multiple_keys(navegador, Keys.ARROW_DOWN, 5)
time.sleep(2)

send_multiple_keys(navegador, Keys.ENTER, 1)
time.sleep(2)

# Confirmar "Boavista"
AplicarFiltro = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter__apply-btn"))
)
AplicarFiltro.click()
time.sleep(4)

NotafilialManaus = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#pdf-dashboard > div > main > div:nth-child(4) > div.ant-col.ant-col-8 > section > div > div > div > div > span")))
NotafilialManaus = NotafilialManaus.text  # Extrair o texto do elemento
time.sleep(2)

NotafilialEqVendasManaus = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#pdf-dashboard > div > main > div:nth-child(8) > div.ant-col.ant-col-8 > section > div > div > div > div > span")))
Nota_filial_EqVendas_Manaus = NotafilialEqVendasManaus.text  # Extrair o texto do elemento
time.sleep(2)

print(f"ATENDIMENTO AO CLIENTE MANAUS: {NotafilialManaus}")
print(f"NOTA FILIAL EQVENDAS MANAUS: {Nota_filial_EqVendas_Manaus}")

# Clicar no botão de filtro de data
filtroData = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#header-filter-button"))
)
filtroData.click()
time.sleep(4)


# Selecionar "Boavista"
desmarcar = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter-container > div:nth-child(4) > div.ant-select.default-select.--block.--medium.filter__branch-select.ant-select-multiple.ant-select-allow-clear.ant-select-show-search > div > div > div:nth-child(1) > span > span.ant-select-selection-item-remove > span > svg > path"))
)
desmarcar.click()
time.sleep(4)

  # Selecionar "Boavista"
filtroConcessionaria = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter-container > div:nth-child(4) > div.ant-select.default-select.--block.--medium.filter__branch-select.ant-select-multiple.ant-select-allow-clear.ant-select-show-search > div > div"))
)
filtroConcessionaria.click()
time.sleep(4)

send_multiple_keys(navegador, Keys.ARROW_DOWN, 6)
time.sleep(2)

send_multiple_keys(navegador, Keys.ENTER, 1)
time.sleep(2)

# Confirmar "Boavista"
AplicarFiltro = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter__apply-btn"))
)
AplicarFiltro.click()
time.sleep(4)

NotafilialNssSenhora = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#pdf-dashboard > div > main > div:nth-child(4) > div.ant-col.ant-col-8 > section > div > div > div > div > span")))
NotafilialNssSenhora = NotafilialNssSenhora.text  # Extrair o texto do elemento
time.sleep(2)

NotafilialEqVendasNssSenhora = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#pdf-dashboard > div > main > div:nth-child(8) > div.ant-col.ant-col-8 > section > div > div > div > div > span")))
Nota_filial_EqVendas_NssSenhora = NotafilialEqVendasNssSenhora.text  # Extrair o texto do elemento
time.sleep(2)

print(f"ATENDIMENTO AO CLIENTE NSS SENHORA: {NotafilialNssSenhora}")
print(f"NOTA FILIAL EQVENDAS NSS SENHORA: {Nota_filial_EqVendas_NssSenhora}")

# Clicar no botão de filtro de data
filtroData = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#header-filter-button"))
)
filtroData.click()
time.sleep(4)


# Selecionar "Boavista"
desmarcar = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter-container > div:nth-child(4) > div.ant-select.default-select.--block.--medium.filter__branch-select.ant-select-multiple.ant-select-allow-clear.ant-select-show-search > div > div > div:nth-child(1) > span > span.ant-select-selection-item-remove > span > svg > path"))
)
desmarcar.click()
time.sleep(4)

  # Selecionar "Boavista"
filtroConcessionaria = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter-container > div:nth-child(4) > div.ant-select.default-select.--block.--medium.filter__branch-select.ant-select-multiple.ant-select-allow-clear.ant-select-show-search > div > div"))
)
filtroConcessionaria.click()
time.sleep(4)

send_multiple_keys(navegador, Keys.ARROW_DOWN, 7)
time.sleep(2)

send_multiple_keys(navegador, Keys.ENTER, 1)
time.sleep(2)

# Confirmar "Boavista"
AplicarFiltro = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter__apply-btn"))
)
AplicarFiltro.click()
time.sleep(4)

NotafilialPalmares = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#pdf-dashboard > div > main > div:nth-child(4) > div.ant-col.ant-col-8 > section > div > div > div > div > span")))
NotafilialPalmares = NotafilialPalmares.text  # Extrair o texto do elemento
time.sleep(2)

NotafilialEqVendasPalmares = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#pdf-dashboard > div > main > div:nth-child(8) > div.ant-col.ant-col-8 > section > div > div > div > div > span")))
Nota_filial_EqVendas_Palmares = NotafilialEqVendasPalmares.text  # Extrair o texto do elemento
time.sleep(2)

print(f"ATENDIMENTO AO CLIENTE PALMARES: {NotafilialPalmares}")
print(f"NOTA FILIAL EQVENDAS PALMARES: {Nota_filial_EqVendas_Palmares}")

# Clicar no botão de filtro de data
filtroData = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#header-filter-button"))
)
filtroData.click()
time.sleep(4)


# Selecionar "Boavista"
desmarcar = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter-container > div:nth-child(4) > div.ant-select.default-select.--block.--medium.filter__branch-select.ant-select-multiple.ant-select-allow-clear.ant-select-show-search > div > div > div:nth-child(1) > span > span.ant-select-selection-item-remove > span > svg > path"))
)
desmarcar.click()
time.sleep(4)

  # Selecionar "Boavista"
filtroConcessionaria = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter-container > div:nth-child(4) > div.ant-select.default-select.--block.--medium.filter__branch-select.ant-select-multiple.ant-select-allow-clear.ant-select-show-search > div > div"))
)
filtroConcessionaria.click()
time.sleep(4)

send_multiple_keys(navegador, Keys.ARROW_DOWN, 8)
time.sleep(2)

send_multiple_keys(navegador, Keys.ENTER, 1)
time.sleep(2)

# Confirmar "Boavista"
AplicarFiltro = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter__apply-btn"))
)
AplicarFiltro.click()
time.sleep(4)

NotafilialPetropolis = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#pdf-dashboard > div > main > div:nth-child(4) > div.ant-col.ant-col-8 > section > div > div > div > div > span")))
NotafilialPetropolis = NotafilialPetropolis.text  # Extrair o texto do elemento
time.sleep(2)

NotafilialEqVendasPetropolis = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#pdf-dashboard > div > main > div:nth-child(8) > div.ant-col.ant-col-8 > section > div > div > div > div > span")))
Nota_filial_EqVendas_Petropolis = NotafilialEqVendasPetropolis.text  # Extrair o texto do elemento
time.sleep(2)

print(f"ATENDIMENTO AO CLIENTE PETROPOLIS: {NotafilialPetropolis}")
print(f"NOTA FILIAL EQVENDAS PETROPOLIS: {Nota_filial_EqVendas_Petropolis}")

# Clicar no botão de filtro de data
filtroData = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#header-filter-button"))
)
filtroData.click()
time.sleep(4)


# Selecionar "Boavista"
desmarcar = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter-container > div:nth-child(4) > div.ant-select.default-select.--block.--medium.filter__branch-select.ant-select-multiple.ant-select-allow-clear.ant-select-show-search > div > div > div:nth-child(1) > span > span.ant-select-selection-item-remove > span > svg > path"))
)
desmarcar.click()
time.sleep(4)

  # Selecionar "Boavista"
filtroConcessionaria = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter-container > div:nth-child(4) > div.ant-select.default-select.--block.--medium.filter__branch-select.ant-select-multiple.ant-select-allow-clear.ant-select-show-search > div > div"))
)
filtroConcessionaria.click()
time.sleep(4)

send_multiple_keys(navegador, Keys.ARROW_DOWN, 9)
time.sleep(2)

send_multiple_keys(navegador, Keys.ENTER, 1)
time.sleep(2)

# Confirmar "Boavista"
AplicarFiltro = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter__apply-btn"))
)
AplicarFiltro.click()
time.sleep(4)

NotafilialSaoGoncalo = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#pdf-dashboard > div > main > div:nth-child(4) > div.ant-col.ant-col-8 > section > div > div > div > div > span")))
NotafilialSaoGoncalo = NotafilialSaoGoncalo.text  # Extrair o texto do elemento
time.sleep(2)

NotafilialEqVendasSaoGoncalo = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#pdf-dashboard > div > main > div:nth-child(8) > div.ant-col.ant-col-8 > section > div > div > div > div > span")))
Nota_filial_EqVendas_SaoGoncalo = NotafilialEqVendasSaoGoncalo.text  # Extrair o texto do elemento
time.sleep(2)

print(f"ATENDIMENTO AO CLIENTE SAO GONCALO: {NotafilialSaoGoncalo}")
print(f"NOTA FILIAL EQVENDAS SAO GONCALO: {Nota_filial_EqVendas_SaoGoncalo}")

# Clicar no botão de filtro de data
filtroData = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#header-filter-button"))
)
filtroData.click()
time.sleep(4)


# Selecionar "Boavista"
desmarcar = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter-container > div:nth-child(4) > div.ant-select.default-select.--block.--medium.filter__branch-select.ant-select-multiple.ant-select-allow-clear.ant-select-show-search > div > div > div:nth-child(1) > span > span.ant-select-selection-item-remove > span > svg > path"))
)
desmarcar.click()
time.sleep(4)

  # Selecionar "Boavista"
filtroConcessionaria = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter-container > div:nth-child(4) > div.ant-select.default-select.--block.--medium.filter__branch-select.ant-select-multiple.ant-select-allow-clear.ant-select-show-search > div > div"))
)
filtroConcessionaria.click()
time.sleep(4)

send_multiple_keys(navegador, Keys.ARROW_DOWN, 20)
time.sleep(2)

send_multiple_keys(navegador, Keys.ENTER, 1)
time.sleep(2)

# Confirmar "Boavista"
AplicarFiltro = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter__apply-btn"))
)
AplicarFiltro.click()
time.sleep(4)

NotafilialSaoLuis = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#pdf-dashboard > div > main > div:nth-child(4) > div.ant-col.ant-col-8 > section > div > div > div > div > span")))
NotafilialSaoLuis = NotafilialSaoLuis.text  # Extrair o texto do elemento
time.sleep(2)

NotafilialEqVendasSaoLuis = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#pdf-dashboard > div > main > div:nth-child(8) > div.ant-col.ant-col-8 > section > div > div > div > div > span")))
Nota_filial_EqVendas_SaoLuis = NotafilialEqVendasSaoLuis.text  # Extrair o texto do elemento
time.sleep(2)

print(f"ATENDIMENTO AO CLIENTE SAO LUIS: {NotafilialSaoLuis}")
print(f"NOTA FILIAL EQVENDAS SAO LUIS: {Nota_filial_EqVendas_SaoLuis}")

# Clicar no botão de filtro de data
filtroData = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#header-filter-button"))
)
filtroData.click()
time.sleep(4)


# Selecionar "Boavista"
desmarcar = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter-container > div:nth-child(4) > div.ant-select.default-select.--block.--medium.filter__branch-select.ant-select-multiple.ant-select-allow-clear.ant-select-show-search > div > div > div:nth-child(1) > span > span.ant-select-selection-item-remove > span > svg > path"))
)
desmarcar.click()
time.sleep(4)

  # Selecionar "Boavista"
filtroConcessionaria = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter-container > div:nth-child(4) > div.ant-select.default-select.--block.--medium.filter__branch-select.ant-select-multiple.ant-select-allow-clear.ant-select-show-search > div > div"))
)
filtroConcessionaria.click()
time.sleep(4)

send_multiple_keys(navegador, Keys.ARROW_DOWN, 11)
time.sleep(2)

send_multiple_keys(navegador, Keys.ENTER, 1)
time.sleep(2)

# Confirmar "Boavista"
AplicarFiltro = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter__apply-btn"))
)
AplicarFiltro.click()
time.sleep(4)

NotafilialTangua = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#pdf-dashboard > div > main > div:nth-child(4) > div.ant-col.ant-col-8 > section > div > div > div > div > span")))
NotafilialTangua = NotafilialTangua.text  # Extrair o texto do elemento
time.sleep(2)

NotafilialEqVendasTangua = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#pdf-dashboard > div > main > div:nth-child(8) > div.ant-col.ant-col-8 > section > div > div > div > div > span")))
Nota_filial_EqVendas_Tangua = NotafilialEqVendasTangua.text  # Extrair o texto do elemento
time.sleep(2)

print(f"ATENDIMENTO AO CLIENTE TANGUA: {NotafilialTangua}")
print(f"NOTA FILIAL EQVENDAS TANGUA: {Nota_filial_EqVendas_Tangua}")

# Clicar no botão de filtro de data
filtroData = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#header-filter-button"))
)
filtroData.click()
time.sleep(4)


# Selecionar "Boavista"
desmarcar = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter-container > div:nth-child(4) > div.ant-select.default-select.--block.--medium.filter__branch-select.ant-select-multiple.ant-select-allow-clear.ant-select-show-search > div > div > div:nth-child(1) > span > span.ant-select-selection-item-remove > span > svg > path"))
)
desmarcar.click()
time.sleep(4)

  # Selecionar "Boavista"
filtroConcessionaria = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter-container > div:nth-child(4) > div.ant-select.default-select.--block.--medium.filter__branch-select.ant-select-multiple.ant-select-allow-clear.ant-select-show-search > div > div"))
)
filtroConcessionaria.click()
time.sleep(4)

send_multiple_keys(navegador, Keys.ARROW_DOWN, 12)
time.sleep(2)

send_multiple_keys(navegador, Keys.ENTER, 1)
time.sleep(2)

# Confirmar "Boavista"
AplicarFiltro = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter__apply-btn"))
)
AplicarFiltro.click()
time.sleep(4)

NotafilialTeresina = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#pdf-dashboard > div > main > div:nth-child(4) > div.ant-col.ant-col-8 > section > div > div > div > div > span")))
NotafilialTeresina = NotafilialTeresina.text  # Extrair o texto do elemento
time.sleep(2)

NotafilialEqVendasTeresina = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#pdf-dashboard > div > main > div:nth-child(8) > div.ant-col.ant-col-8 > section > div > div > div > div > span")))
Nota_filial_EqVendas_Teresina = NotafilialEqVendasTeresina.text  # Extrair o texto do elemento
time.sleep(2)

print(f"ATENDIMENTO AO CLIENTE TERESINA: {NotafilialTeresina}")
print(f"NOTA FILIAL EQVENDAS TERESINA: {Nota_filial_EqVendas_Teresina}")

# Clicar no botão de filtro de data
filtroData = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#header-filter-button"))
)
filtroData.click()
time.sleep(4)


# Selecionar "Boavista"
desmarcar = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter-container > div:nth-child(4) > div.ant-select.default-select.--block.--medium.filter__branch-select.ant-select-multiple.ant-select-allow-clear.ant-select-show-search > div > div > div:nth-child(1) > span > span.ant-select-selection-item-remove > span > svg > path"))
)
desmarcar.click()
time.sleep(4)

  # Selecionar "Boavista"
filtroConcessionaria = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter-container > div:nth-child(4) > div.ant-select.default-select.--block.--medium.filter__branch-select.ant-select-multiple.ant-select-allow-clear.ant-select-show-search > div > div"))
)
filtroConcessionaria.click()
time.sleep(4)

send_multiple_keys(navegador, Keys.ARROW_DOWN, 12)
time.sleep(2)

send_multiple_keys(navegador, Keys.ENTER, 1)
time.sleep(2)

# Confirmar "Boavista"
AplicarFiltro = WebDriverWait(navegador, 20).until(
EC.element_to_be_clickable((By.CSS_SELECTOR, "#filter__apply-btn"))
)
AplicarFiltro.click()
time.sleep(4)

NotafilialUrucui = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#pdf-dashboard > div > main > div:nth-child(4) > div.ant-col.ant-col-8 > section > div > div > div > div > span")))
NotafilialUrucui = NotafilialUrucui.text  # Extrair o texto do elemento
time.sleep(2)

NotafilialEqVendasUrucui = WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#pdf-dashboard > div > main > div:nth-child(8) > div.ant-col.ant-col-8 > section > div > div > div > div > span")))
Nota_filial_EqVendas_Urucui = NotafilialEqVendasUrucui.text  # Extrair o texto do elemento
time.sleep(2)

print(f"ATENDIMENTO AO CLIENTE URUCUI: {NotafilialUrucui}")
print(f"NOTA FILIAL EQVENDAS URUCUI: {Nota_filial_EqVendas_Urucui}")




# Dicionário com os dados das empresas e suas respectivas notas
dados = {

    "Nacional": (Nota_NacionalAtn, Nota_Nacional_Equipe_Vendas),
    "Boavista": (NotafilialBoaVista, Nota_filial_EqVendas_Boavista),
    "Brasilia": (NotafilialBrasilia, Nota_filial_EqVendas_Brasilia),
    "Campos": (NotafilialCampos, Nota_filial_EqVendas_Campos),
    "Floriano": (NotafilialFloriano, Nota_filial_EqVendas_Floriano),
    "Manaus": (NotafilialManaus, Nota_filial_EqVendas_Manaus),
    "NossaSenhora": (NotafilialNssSenhora, Nota_filial_EqVendas_NssSenhora),
    "SãoLuis": (NotafilialSaoLuis, Nota_filial_EqVendas_SaoLuis),
    "Teresina": (NotafilialTeresina, Nota_filial_EqVendas_Teresina),
    "São Goncalo": (NotafilialSaoGoncalo, Nota_filial_EqVendas_SaoGoncalo),
    "Tangua": (NotafilialTangua, Nota_filial_EqVendas_Tangua),
    "Petropolis": (NotafilialPetropolis, Nota_filial_EqVendas_Petropolis),
    "Luziana": (NotafilialLuziana, Nota_filial_EqVendas_Luziana),
    "Palmares": (NotafilialPalmares, Nota_filial_EqVendas_Palmares),
    "Urucui": (NotafilialUrucui, Nota_filial_EqVendas_Urucui)
}


df = pd.DataFrame(dados, index=["Atendimento_ao_Cliente", "Nota_filial_EqVendas"]).transpose()

# Renomeando as colunas
df.columns = ["Nota_filial_Anual", "Nota_filial_EqVendas_Anual"]

# Adicionando a coluna de Empresa
df["Empresa"] = df.index

# Reorganizando as colunas
df = df[["Empresa", "Nota_filial_Anual", "Nota_filial_EqVendas_Anual"]]

# Removendo a coluna sem nome (índice neste caso)
df.reset_index(drop=True, inplace=True)

# Adicionando a coluna "data_atualizacao" com a data atual
df['data_atualizacao'] = date.today()

# Exibindo o DataFrame
print(df)

##########################################################################
#conexâo ao banco de dados

print("Conectando ao banco de dados...")
user = 'rpa_bi'
password = 'Rp@_B&_P@rvi'
host = '10.0.10.243'
port = '54949'
database = 'stage'

params = urllib.parse.quote_plus(
    f'DRIVER=ODBC Driver 17 for SQL Server;SERVER={host},{port};DATABASE={database};UID={user};PWD={password}')
connection_str = f'mssql+pyodbc:///?odbc_connect={params}'
engine = create_engine(connection_str)
table_name = "AtendimentoCli_Ano_IndeCX"

with engine.connect() as connection:
    df.to_sql(table_name, con=connection, if_exists='replace', index=False)

print(f"Dados inseridos com sucesso na tabela '{table_name}'!")
print("Fechando o navegador...")
navegador.quit()