from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import os, time
from datetime import date
import datetime
from random import randrange
import keyboard


opt = webdriver.ChromeOptions() 
opt.add_experimental_option("excludeSwitches", ["enable-automation"])
opt.add_experimental_option('useAutomationExtension', False)
opt.add_argument("--window-size=1300,620")

driver = webdriver.Chrome(options=opt, executable_path= r'C:\Users\Take4\Documents\Saves\chromedriver_win32\chromedriver.exe')

driver.get("https://homolog.farmazon.com.br")

mail_addressDjango = "-"
passwordDjango = "-"

WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.NAME, "username"))).send_keys(mail_addressDjango)
WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.NAME, "password"))).send_keys(passwordDjango)
WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div/form[2]/button'))).click()

lista646 = []
lista647 = []
lista648 = []

def adicionarProduto1(lista):
    
    while len(lista) <= 9:
        produtoAleatorio = randrange(47557)
        
        driver.get('https://homolog.farmazon.com.br/api/products/'+str(produtoAleatorio)+'/?format=json')
        
        with open('vitrine13.txt', 'w', encoding = 'utf-16') as f:
            f.write(driver.page_source)
            
        arquivo69 = (r'C:\Users\Take4\Desktop\KPI\vitrine13.txt')
        
        #Arquivo TXT salvo é lido para ver quantas repeticoes do "create time" + ano + mes existem.
        f69 = open(arquivo69, encoding = 'utf-16')
        file_contents69 = f69.read()
        
        product1 = file_contents69.count('produto-sem-imagem')
        product2 = file_contents69.count('"price":0.0')
        product3 = file_contents69.count('default')
            
        print (produtoAleatorio)
                
        if product1 == 1 or product2 == 1 or product3 == 1:
            pass
        elif product1 == 0 and product2 == 0 and product3 == 0:
            lista.append(str(produtoAleatorio))
        else:
            pass
        
        print (lista)

def adicionarProduto2 (lista, categoria):
    index = 0
    while index < len(lista):
        driver.get('https://homolog.farmazon.com.br/admin/api/product/'+lista[index])
        
        keyboard.press('ctrl')
        WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div[3]/div/form/div/fieldset/div[1]/div/div[1]/select/option['+str(categoria)+']'))).click()
        time.sleep(3)
        
        vazio1 = driver.find_element_by_xpath("/html/body/div/div[3]/div/form/div/fieldset/div[7]/div/div[1]/select/option[1969]").is_selected()
        if vazio1 == True:
            pass
        elif vazio1 == False:
            WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div[3]/div/form/div/fieldset/div[7]/div/div[1]/select/option[1969]'))).click()
        time.sleep(3)
        
        vazio2 = driver.find_element_by_xpath("html/body/div/div[3]/div/form/div/fieldset/div[12]/div/div[1]/select/option[118]").is_selected()
        if vazio2 == True:
            pass
        elif vazio2 == False:
            WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div[3]/div/form/div/fieldset/div[12]/div/div[1]/select/option[118]'))).click()
        time.sleep(3)
        
        keyboard.release('ctrl')
        
        WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div[3]/div/form/div/div/input[3]'))).click()
        index += 1


adicionarProduto1(lista646)
adicionarProduto1(lista647)
adicionarProduto1(lista648)

time.sleep(10)

adicionarProduto2(lista646, 646)
adicionarProduto2(lista647, 647)
adicionarProduto2(lista648, 648)
    
driver.close()
     
    
    
    

    
    
    
    
    
    



