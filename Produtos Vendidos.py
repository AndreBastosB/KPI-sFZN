import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import os

# AUTOMAÇÃO DOS SEGUINTES KPI'S NESTE CÓDIGO:
    # - PEDIDOS CONCLUÍDOS
    # - PEDIDOS CANCELADOS
    # - PEDIDOS TESTE

planilhaEditada = 'Teste1'

ano = '2021'
mes = '01'

diainicial = "01"

if mes == "01" or mes =="03" or mes == "05" or mes == "07" or mes == "08" or mes == "10" or mes == "12":
    diafinal = "31"
if mes == "04" or mes == "06" or mes == "09" or mes == "11":
    diafinal = "30"
elif mes == "02":
    diafinal = "28"
    
#AUTENTICAÇÃO AO GOOGLE SHEET - AUTENTICAÇÃO AO GOOGLE SHEET - AUTENTICAÇÃO AO GOOGLE SHEET - AUTENTICAÇÃO AO GOOGLE SHEET
#AUTENTICAÇÃO AO GOOGLE SHEET - AUTENTICAÇÃO AO GOOGLE SHEET - AUTENTICAÇÃO AO GOOGLE SHEET - AUTENTICAÇÃO AO GOOGLE SHEET - 
scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)

client = gspread.authorize(creds)

#Seleciona o arquivo que estarei visualizando / editando.
sheet = client.open(planilhaEditada).sheet1
    
opt = webdriver.ChromeOptions() 
opt.add_argument("--auto-open-devtools-for-tabs")
opt.add_experimental_option("excludeSwitches", ["enable-automation"])
opt.add_experimental_option('useAutomationExtension', False)
opt.add_argument("--window-size=1034,620")

#Seleciona o local de instalação do webdriver e as opções listadas acima.
driver = webdriver.Chrome(options=opt, executable_path= r'C:\Users\Take4\Documents\Saves\chromedriver_win32\chromedriver.exe')


#CONTABILIZAR PEDIDOS DJANGO - CONTABILIZAR PEDIDOS DJANGO - CONTABILIZAR PEDIDOS DJANGO - CONTABILIZAR PEDIDOS DJANGO
driver.get("https://homolog.farmazon.com.br")

mail_addressDjango = "-"
passwordDjango = "-"

#Autenticação no Django.
WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.NAME, "username"))).send_keys(mail_addressDjango)
WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.NAME, "password"))).send_keys(passwordDjango)
WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div/form[2]/button'))).click()

driver.get ('https://homolog.farmazon.com.br/api/orders/?format=json&limit=1500')

#Primeira pagina de novos usuarios é salva como JSON no arquivo TXT
with open('orders.txt', 'w') as f:
    f.write(driver.page_source)

arquivo1 = (r'C:\Users\Take4\Desktop\KPI\\orders.txt')

#Arquivo TXT salvo é lido para ver quantas repeticoes do "create time" + ano + mes existem.
f1 = open(arquivo1)
file_contents1 = f1.read()

firstOrder = file_contents1.split('count":',1)[1]
lastOrder = firstOrder.split(',"next":', 1)[0]

janeiro = file_contents1.count(',"create_time":"'+ano+'-01')
fevereiro = file_contents1.count(',"create_time":"'+ano+'-02')
março = file_contents1.count(',"create_time":"'+ano+'-03')
abril = file_contents1.count(',"create_time":"'+ano+'-04')
maio = file_contents1.count(',"create_time":"'+ano+'-05')
junho = file_contents1.count(',"create_time":"'+ano+'-06')
julho = file_contents1.count(',"create_time":"'+ano+'-07')
agosto = file_contents1.count(',"create_time":"'+ano+'-08')
setembro = file_contents1.count(',"create_time":"'+ano+'-09')
outubro = file_contents1.count(',"create_time":"'+ano+'-10')
novembro = file_contents1.count(',"create_time":"'+ano+'-11')
dezembro = file_contents1.count(',"create_time":"'+ano+'-12')

if mes == '01':
    link = str(int(lastOrder)-(janeiro + fevereiro + março + abril + maio + junho + julho + agosto + setembro + outubro + novembro + dezembro -1))
elif mes == '02':
    link = str(int(lastOrder)-(fevereiro + março + abril + maio + junho + julho + agosto + setembro + outubro + novembro + dezembro -1))
elif mes == '03':
    link = str(int(lastOrder)-(março + abril + maio + junho + julho + agosto + setembro + outubro + novembro + dezembro -1))
elif mes == '04':
    link = str(int(lastOrder)-(abril + maio + junho + julho + agosto + setembro + outubro + novembro + dezembro -1))
elif mes == '05':
    link = str(int(lastOrder)-(maio + junho + julho + agosto + setembro + outubro + novembro + dezembro -1))
elif mes == '06':
    link = str(int(lastOrder)-(junho + julho + agosto + setembro + outubro + novembro + dezembro -1))
elif mes == '07':
    link = str(int(lastOrder)-(julho + agosto + setembro + outubro + novembro + dezembro -1))
elif mes == '08':
    link = str(int(lastOrder)-(agosto + setembro + outubro + novembro + dezembro -1))
elif mes == '09':
    link = str(int(lastOrder)-(setembro + outubro + novembro + dezembro -1))
elif mes == '10':
    link = str(int(lastOrder)-(outubro + novembro + dezembro -1))
elif mes == '11':
    link = str(int(lastOrder)-(novembro + dezembro -1))
elif mes == '12':
    link = str(int(lastOrder)-(dezembro -1))

f1.close()

idOrder = int(link)


while idOrder <= int(lastOrder):
    driver.get('https://homolog.farmazon.com.br/api/orders/'+str(idOrder)+'/?format=json')
    
    with open('pedidos.txt', 'w') as f:
        f.write(driver.page_source)
        
    arquivos3 = (r'C:\Users\Take4\Desktop\KPI\\pedidos.txt')
    f3 = open(arquivos3)
    file_contents3 = f3.read()
    
    productList = file_contents3.split(',"products"', 1)[1]
    productfinalList = productList.split('}],"store"', 1)[0]
    
    quantidadeDeProdutos = productfinalList.count('product_id":')
    
    prodcutAmount = 1
    
    while prodcutAmount <= quantidadeDeProdutos:
        product1 = productfinalList.split(',"name":"', 50)[prodcutAmount]
        product11 = product1.split('","scientific_name"', 1)[0]
        print ('- ' + (product11))
        prodcutAmount += 1
    
    
    idOrder += 1
    
    
driver.close()
f3.close()
os.remove('orders.txt')
os.remove('pedidos.txt')










