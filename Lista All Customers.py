from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time

#CODIGO CRIADO PARA PUXAR TODOS OS CADASTROS REALIZADOS NO APP FARMAZON E APLICA-LOS NA PLANILHA ESCOLHIDA ABAIXO.

planilhaEditada = "KPI's Interno"

scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)

client = gspread.authorize(creds)

#Seleciona o arquivo que estarei visualizando / editando.
sheet = client.open(planilhaEditada).worksheet("Cadastros") 
sheet2 = client.open(planilhaEditada).worksheet("Clientes x Pedidos") 

opt = webdriver.ChromeOptions() 
opt.add_experimental_option("excludeSwitches", ["enable-automation"])
opt.add_experimental_option('useAutomationExtension', False)
opt.add_argument("--window-size=1034,620")

driver = webdriver.Chrome(options=opt, executable_path= r'C:\Users\Take4\Documents\Saves\chromedriver_win32\chromedriver.exe')

driver.get("https://homolog.farmazon.com.br")

mail_addressDjango = "-"
passwordDjango = "-"

#Autenticação no Django.
WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.NAME, "username"))).send_keys(mail_addressDjango)
WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.NAME, "password"))).send_keys(passwordDjango)
WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div/form[2]/button'))).click()


def next_available_row(worksheet):
    str_list = list(filter(None, worksheet.col_values(1)))
    return str(len(str_list)+1)

#Proxima Linha disponivel na Planilha = linha
linha = next_available_row(sheet)

linha1 = (int(linha) - 1)

clientID1 = sheet.cell(linha1, 1).value

#Proximo cadastro que nao está na planilha.
clientID = (int(clientID1) +1)

driver.get('https://homolog.farmazon.com.br/api/customers/?format=json&limit=1')

with open('clientes_total.txt', 'w') as f:
    f.write(driver.page_source)

arquivo253 = (r'C:\Users\Take4\Desktop\KPI\\clientes_total.txt')

#Arquivo TXT salvo é lido para ver quantas repeticoes do "create time" + ano + mes existem.
f253 = open(arquivo253)
file_contents253 = f253.read()

file_contents2531 = file_contents253.split(',"id":', 1)[1]
file_contents2532 = file_contents2531.split('}]}', 1)[0]

##ULTIMO CLIENTE A SE CADASTRAR = cliente_final
cliente_final = file_contents2532

print (linha)
print (clientID)
print (cliente_final)

linha = int(linha)

time.sleep(10)

while int(clientID) <= int(cliente_final):
    driver.get ('https://homolog.farmazon.com.br/api/customers/'+str(clientID)+'/?format=json')
    
    with open('clientes.txt', 'w', encoding='utf-16') as f:
        f.write(driver.page_source)
    
    arquivo12 = (r'C:\Users\Take4\Desktop\KPI\\clientes.txt')
    
    f12 = open(arquivo12, encoding='utf-16')
    file_contents12 = f12.read()
    
    contagem1 = file_contents12.count('{"detail":"Não encontrado."}')
    
    if contagem1 == 1:
        pass
    
    else:
        clienteId = file_contents12.split('json","id":',1)[1]
        finalClienteId = clienteId.split('}', 1)[0]
        
        userID = file_contents12.split('https://homolog.farmazon.com.br/api/users/',1)[1]
        finalUserID = userID.split('/?format=',1)[0]
        
        clientEmail = file_contents12.split('"email":"',1)[1]
        finalClientEmail = clientEmail.split('",',1)[0]
        
        clientTime = file_contents12.split('"create_time":"',1)[1]
        finalClientTime = clientTime.split('T',1)[0]
        
        clientName = file_contents12.split('"name":"',1)[1]
        finalClientName = clientName.split('","telephone',1)[0]

        
        clientPhone = file_contents12.split('"telephone":',1)[1]
        finalPhone = clientPhone.split(',"',1)[0]
        
        sheet.update_cell(linha, 1, finalClienteId)
        sheet.update_cell(linha, 5, finalClientEmail)
        sheet.update_cell(linha, 3, finalClientTime)
        sheet.update_cell(linha, 2, finalClientName)
        sheet.update_cell(linha, 4, finalPhone)
        
        sheet2.update_cell(linha, 1, finalClienteId)
        sheet2.update_cell(linha, 2, finalClientName)

        linha += 1
    
    clientID += 1
    time.sleep(15)

#f12.close()
driver.close()

