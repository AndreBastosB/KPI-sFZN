import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import os, time
from datetime import date
import datetime

planilhaEditada = "KPI's Interno"
planilhaBase = 'Relatorio de Pedidos - Oficiosa (2020/2021)'

#ATENÇÃO: INSERIR NA VARIAVEL LINHA A LINHA QUE CORRESPONDE AO PRIMEIRO PEDIDO DO MES O QUAL ESTA SENDO PESQUISADO.
#COMANDO ADICIONADO PARA A CONTAGEM DOS PAGAMENTOS FEITOS POR PIX.
linha = 2

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
    
#AUTENTICAÇÃO AO GOOGLE SHEET - AUTENTICAÇÃO AO GOOGLE SHEET - AUTENTICAÇÃO AO GOOGLE SHEET - AUTENTICAÇÃO AO GOOGLE SHEET
#AUTENTICAÇÃO AO GOOGLE SHEET - AUTENTICAÇÃO AO GOOGLE SHEET - AUTENTICAÇÃO AO GOOGLE SHEET - AUTENTICAÇÃO AO GOOGLE SHEET - 
scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)

client = gspread.authorize(creds)

#Seleciona o arquivo que estarei visualizando / editando.
sheet2 = client.open(planilhaEditada).worksheet("KPI's")
sheet = client.open(planilhaBase).worksheet("GERAL (a partir de 3 set.20)") 
    
opt = webdriver.ChromeOptions() 
opt.add_argument("--auto-open-devtools-for-tabs")
opt.add_experimental_option("excludeSwitches", ["enable-automation"])
opt.add_experimental_option('useAutomationExtension', False)
opt.add_argument("--window-size=1034,620")

#Seleciona o local de instalação do webdriver e as opções listadas acima.
driver = webdriver.Chrome(options=opt, executable_path= r'C:\Users\Take4\Documents\Saves\chromedriver_win32\chromedriver.exe')

driver.get("https://merchants.payulatam.com/m/853003/a/860542/reports/salesReport?repFil=ps%3D10%26p%3D1%26fd%3D"+ano+'%252F'+mes+'%252F'+diainicial+"%26td%3D"+ano+"%252F"+mes+"%252F"+diafinal+"%26os%3DCAPTURED")

loginPayU = "-"
passwordPayU = "-"

WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.NAME, "username"))).send_keys(loginPayU)
WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.NAME, "password"))).send_keys(passwordPayU)
WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div/div/main/div/div[3]/div/div/form/div/div/center/button/span"))).click()

payUPurchases2 = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div/div/main/div/section/div[4]/div[2]/div[3]/ng-include[1]/div/div[1]/div/div/span[2]'))).text

payUPurchases = payUPurchases2.split('Registros de ', 1)[1]

driver.get("https://homolog.farmazon.com.br")

mail_addressDjango = "-"
passwordDjango = "-"

WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.NAME, "username"))).send_keys(mail_addressDjango)
WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.NAME, "password"))).send_keys(passwordDjango)
WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div/form[2]/button'))).click()

driver.get ('https://homolog.farmazon.com.br/api/orders/?format=json&limit=5000')

#Primeira pagina de novos usuarios é salva como JSON no arquivo TXT
with open('orders.txt', 'w') as f:
    f.write(driver.page_source)

arquivo1 = (r'C:\Users\Take4\Desktop\KPI\\orders.txt')

#Arquivo TXT salvo é lido para ver quantas repeticoes do "create time" + ano + mes existem.
f1 = open(arquivo1)
file_contents1 = f1.read()

firstOrder = file_contents1.split('","create_time":"' + ano + "-" + mes, 1)[1]
lastOrder2 = firstOrder.split(',"url"', 1)[0]
lastOrder = lastOrder2.split(',{"id":', 1)[1]

ultimoPedidoDoMes = int(lastOrder) + 1

f1.close()

pedidosTestes = 0
pedidosCancelados = 0
pedidosReais = 0

def lendoPedidos(pedidoInicial):
    driver.get('https://homolog.farmazon.com.br/api/orders/'+ str(pedidoInicial) +'/?format=json')
    with open('pedidos.txt', 'w', encoding = 'utf-16') as f:
        f.write(driver.page_source)
    arquivos2 = (r'C:\Users\Take4\Desktop\KPI\\pedidos.txt')
    f2 = open(arquivos2, encoding = 'utf-16')
    file_contents2 = f2.read()
    lendoData = file_contents2.count(',"create_time":"' + ano + '-' + mes)
    lendoData2 = file_contents2.count('},"create_time":"' + ano + '-' + mes)
    if lendoData == 1 or lendoData2 == 1:
        return ('true')
    else:
        return ('false')
    f2.close()
        
    

while lendoPedidos(ultimoPedidoDoMes) == 'true':
    driver.get('https://homolog.farmazon.com.br/api/orders/'+ str(ultimoPedidoDoMes) +'/?format=json')
    with open('pedidos2.txt', 'w', encoding = 'utf-16') as f:
        f.write(driver.page_source)
    arquivos3 = (r'C:\Users\Take4\Desktop\KPI\\pedidos2.txt')
    f3 = open(arquivos3, encoding = 'utf-16')
    file_contents3 = f3.read()
    customer1IDCount = file_contents3.count('"customer_id":"3"')
    customer2IDCount = file_contents3.count('"customer_id":"1023"')
    customer3IDCount = file_contents3.count('"customer_id":"1112"')
    customer4IDCount = file_contents3.count('"customer_id":"1"')
    customer5IDCount = file_contents3.count('"customer_id":"932"')
    customer6IDCount = file_contents3.count('"customer_id":"1167"')
    customer7IDCount = file_contents3.count('"customer_id":"5"')
    cancelCount = file_contents3.count('"cancelado"')
    paymentCount = file_contents3.count('"payments":[]')
    realCount1 = file_contents3.count('},"create_time":"'+ano+'-'+mes)
    realCount2 = file_contents3.count('l,"create_time":"'+ano+'-'+mes)
    contagemEntregue = file_contents3.count('entregue')
    pagamentoContagem = file_contents3.count('"operator_payment_id":')
    if realCount1 == 1 or realCount2 == 1:
        if customer1IDCount == 1 or customer2IDCount == 1 or customer3IDCount == 1 or customer4IDCount == 1 or customer5IDCount == 1 or customer6IDCount == 1 or customer7IDCount == 1:
            pedidosTestes += 1
            print ('pedido teste ' + str(ultimoPedidoDoMes))
        elif cancelCount == 2:
            pedidosCancelados += 1
            print ('pedido cancelado ' + str(ultimoPedidoDoMes))
        elif contagemEntregue == 2 or pagamentoContagem == 1:
            pedidosReais += 1
            print ('pedido real ' + str(ultimoPedidoDoMes))
        else:
            pedidosCancelados += 1
            print ("Verificar: " + str(ultimoPedidoDoMes))
    else:

        f3.close()
        break
    
    ultimoPedidoDoMes -= 1
    
driver.close()
os.remove('orders.txt')
os.remove('pedidos.txt')


pedidosTotais = int(pedidosReais) + int(payUPurchases)

print ('Testes: ' + str(pedidosTestes))
print ('Cancelados: ' + str(pedidosCancelados))
print ('Concluidos: ' + str(pedidosTotais))

data = mes + '/' + ano

while data in sheet.cell(linha, 3).value:
    if "Pix" in sheet.cell(linha, 2).value:
        pedidosTotais += 1
        print (linha)
    else:
        pass
    linha += 1
    time.sleep(5)
    
#Linha correspondentes no KPI.
linhaPedidosCancelados = 11
linhaPedidosTestes = 12
linhaPedidosConcluidos = 7

#Coluna do mês correspondente.
if ano == "2019":
    Mes = int(mes) - 5
if ano == "2020":
    Mes = int(mes) + 7
if ano == "2021":
    Mes = int(mes) + 19
elif ano == "2022":
    Mes = int(mes) + 31

#Atualiza na tabela do Google Sheet os valores.
sheet2.update_cell (linhaPedidosConcluidos, Mes, pedidosTotais)
sheet2.update_cell (linhaPedidosCancelados, Mes, pedidosCancelados)
sheet2.update_cell (linhaPedidosTestes, Mes, pedidosTestes)