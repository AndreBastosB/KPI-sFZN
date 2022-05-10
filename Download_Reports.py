from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import os, time
from datetime import date
import datetime

#CRIADO POR DECOGOL
#CODIGO CRIADO PARA DOWNLOAD DOS RELATORIOS DOS KPI'S ABAIXO PARA PREENCHIMENTO PORTERIOR (EM OUTRO CODIGO) 
#NA PLANILHA DESEJADA POSTERIORMENTE DESEJADA.
    # - NOVAS INSTALAÇÕES.
    # - USUÁRIOS ATIVOS.
    # - DESINSTALAÇÕES.
    
data = datetime.datetime.now()    

ano = str(data.year)

mes = str(data.month)

if (len(mes)) == 1:
    mes = '0' + mes
else:
    pass

diainicial = '01'

diafinal1 = (data.day)
diafinal2 = int(diafinal1) - 1
diafinal = str(diafinal2)

if (len(diafinal)) == 1:
    diafinal = '0' + diafinal

armazenamentoRelatorios = r'C:\Users\Take4\Desktop\KPI\\'+mes+ano+'\\'

if not os.path.exists(armazenamentoRelatorios):
    os.makedirs(armazenamentoRelatorios)
    
#Passo 1: Colocar as configurações do selenium webdriver da maneira correta a importar os relatorios.
opt = webdriver.ChromeOptions() 
opt.add_argument("--auto-open-devtools-for-tabs")
opt.add_experimental_option("excludeSwitches", ["enable-automation"])
opt.add_experimental_option('useAutomationExtension', False)
opt.add_argument("--window-size=1034,620")

#Seleciona o local de instalação do webdriver e as opções listadas acima.
driver = webdriver.Chrome(options=opt, executable_path= r'C:\Users\Take4\Documents\Saves\chromedriver_win32\chromedriver.exe')

#Abre o Chrome webdriver na página de métricas da Apple.
driver.get ("https://appstoreconnect.apple.com/login?targetUrl=%2Fanalytics&authResult=FAILED")

#Área de login na conta Andre Bastos - ADM do app Farmazon.
#Aguarda-se 30s para carregar a pagina e preenche com o e-mail.
mail_address = "-"
password_mail = "-"

#Autenticação no Apple Analytics, aguardando o campo estar disponivel para ser preenchido.
WebDriverWait(driver, 120).until(EC.frame_to_be_available_and_switch_to_it((By.ID,"aid-auth-widget-iFrame")))
WebDriverWait(driver, 160).until(EC.element_to_be_clickable((By.ID, "account_name_text_field"))).send_keys(mail_address)
WebDriverWait(driver, 160).until(EC.element_to_be_clickable((By.ID, 'sign-in'))).click()
WebDriverWait(driver, 160).until(EC.element_to_be_clickable((By.ID, 'password_text_field'))).send_keys(password_mail)
WebDriverWait(driver, 160).until(EC.element_to_be_clickable((By.ID, 'sign-in'))).click()

driver.minimize_window()

secretkey = input ("Insira cod.: ")

driver.maximize_window()
driver.set_window_size(1079, 532)

#Insere o código de 6 digitos no local correto.
codigo = driver.find_element_by_id('char0')
codigo.send_keys(secretkey[0])
codigo = driver.find_element_by_id('char1')
codigo.send_keys(secretkey[1])
codigo = driver.find_element_by_id('char2')
codigo.send_keys(secretkey[2])
codigo = driver.find_element_by_id('char3')
codigo.send_keys(secretkey[3])
codigo = driver.find_element_by_id('char4')
codigo.send_keys(secretkey[4])
codigo = driver.find_element_by_id('char5')
codigo.send_keys(secretkey[5])

#Autoriza o Selenium como usuario confiável.
WebDriverWait(driver, 3000).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[4]/div/div/div[1]/div[3]/div/button[3]'))).click()

time.sleep(10)

#Carrega a primeira página para baixar o relatório de Instalações.
driver.get("https://appstoreconnect.apple.com/analytics/app/r:"+ano + mes + diainicial + ':' + ano + mes+diafinal + "/1453428847/metrics?annotationsVisible=true&chartType=singleaxis&measureKey=installs&zoomType=day")

#Pressiona o botão de download do relatório de Usuários Ativos.
WebDriverWait(driver, 160).until(EC.element_to_be_clickable((By.ID, 'Icons/More'))).click()

WebDriverWait(driver, 160).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/ul/form/li/button'))).click()

time.sleep(3)

#Carrega a página para baixar o relatório de Usuários Ativos.
driver.get("https://appstoreconnect.apple.com/analytics/app/r:"+ano + mes + diainicial + ':' + ano + mes+diafinal + "/1453428847/metrics?annotationsVisible=true&chartType=singleaxis&measureKey=activeDevices&zoomType=day")

#Pressiona o botão de download do relatório de Usuários Ativos.
WebDriverWait(driver, 160).until(EC.element_to_be_clickable((By.ID, 'Icons/More'))).click()

WebDriverWait(driver, 160).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/ul/form/li/button'))).click()

time.sleep(3)

#Carrega a página para baixar o relatório de Desinstalações.
driver.get("https://appstoreconnect.apple.com/analytics/app/r:"+ano + mes + diainicial + ':' + ano + mes+diafinal + "/1453428847/metrics?annotationsVisible=true&chartType=singleaxis&measureKey=uninstalls&zoomType=day")

#Pressiona o botão de download do relatório de Usuários Ativos.
WebDriverWait(driver, 160).until(EC.element_to_be_clickable((By.ID, 'Icons/More'))).click()

WebDriverWait(driver, 160).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div/ul/form/li/button'))).click()

time.sleep(3)

##IMPORT DE RELATÓRIOS GOOGLE - IMPORT DE RELATÓRIOS GOOGLE - IMPORT DE RELATÓRIOS GOOGLE - IMPORT DE RELATÓRIOS GOOGLE
##IMPORT DE RELATÓRIOS GOOGLE - IMPORT DE RELATÓRIOS GOOGLE - IMPORT DE RELATÓRIOS GOOGLE - IMPORT DE RELATÓRIOS GOOGLE

#Abre o Chrome webdriver na página de métricas do Google.
driver.get("https://play.google.com/console/u/0/developers/8275367680006099847/app/4973135647988843819/statistics?metrics=ACTIVE_USERS-ALL-UNIQUE-PER_INTERVAL-DAY&dimension=COUNTRY&dimensionValues=OVERALL&dateRange="+ano+'_'+mes+'_'+diainicial+'-'+ano+'_'+mes+'_'+diafinal+"&growthRateDimensionValue=OVERALL&peersetKey=3%3A36b401301830ca59&tab=APP_STATISTICS&ctpMetric=DAU_MAU-ACQUISITION_UNSPECIFIED-COUNT_UNSPECIFIED-CALCULATION_UNSPECIFIED-DAY&ctpDimension=COUNTRY&ctpDimensionValue=OVERALL&ctpPeersetKey=3%3A36b401301830ca59")


#Área de login na conta Andre Bastos - ADM do app Farmazon.
mail_addressGoogle = "-"
password_mailGoogle = "-"

#Preenchimento das credenciais Google no Selenium.
WebDriverWait(driver, 160).until(EC.element_to_be_clickable((By.NAME, "identifier"))).send_keys(mail_addressGoogle)
WebDriverWait(driver, 160).until(EC.element_to_be_clickable((By.ID, 'identifierNext'))).click()

WebDriverWait(driver, 160).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="password"]/div[1]/div/div[1]/input'))).send_keys(password_mailGoogle)
WebDriverWait(driver, 160).until(EC.element_to_be_clickable((By.ID, 'passwordNext'))).click()

#Baixa o primeiro relatório Google.
WebDriverWait(driver, 160).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/root/console-chrome/div/div/div/div/div[1]/page-router-outlet/page-wrapper/div/statistics-page/app-statistics/configure-report/console-block-1-column[1]/div/div/configure-report-header/console-header/div/div/div[1]/div[2]/div/console-button-set/console-button-menu/button-wrapper/button/material-icon/i'))).click()
time.sleep(3)
WebDriverWait(driver, 160).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[5]/div/div/div[2]/div[2]/material-list/div/material-select-item[1]/div/div/div'))).click()
time.sleep(3)
WebDriverWait(driver, 160).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[5]/div/div/div[2]/div[2]/material-list/div/material-select-item[1]'))).click()
time.sleep(3)

#Download do segundo relatório
driver.get("https://storage.cloud.google.com/pubsite_prod_8275367680006099847/stats/installs/installs_br.com.farmazon.client_"+ano+mes+"_overview.csv")

time.sleep(5)

driver.close()