from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import os, time
from datetime import date
import datetime
from random import randrange
import keyboard


def zerandoCategoria (categoryID):
    opt = webdriver.ChromeOptions() 
    opt.add_experimental_option("excludeSwitches", ["enable-automation"])
    opt.add_experimental_option('useAutomationExtension', False)
    opt.add_argument("--window-size=1300,620")
    
    driver = webdriver.Chrome(options=opt, executable_path= r'C:\Users\Take4\Documents\Saves\chromedriver_win32\chromedriver.exe')
    
    driver2 = webdriver.Chrome(options=opt, executable_path= r'C:\Users\Take4\Documents\Saves\chromedriver_win32\chromedriver.exe')
    
    driver.get("https://homolog.farmazon.com.br")
    
    mail_addressDjango = "-"
    passwordDjango = "-"
    
    WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.NAME, "username"))).send_keys(mail_addressDjango)
    WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.NAME, "password"))).send_keys(passwordDjango)
    WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div/form[2]/button'))).click()
    
    driver.get('https://homolog.farmazon.com.br/api/products/?anvisa=&drug=&format=json&id__in=&image_url=&manufacturer=&max_price=&min_price=&name=&product_category='+categoryID+'&scientific_name=')
    
    with open('vitrine1.txt', 'w', encoding = 'utf-16') as f:
        f.write(driver.page_source)
    
    arquivo67 = (r'C:\Users\Take4\Desktop\KPI\vitrine1.txt')
    
    #Arquivo TXT salvo é lido para ver quantas repeticoes do "create time" + ano + mes existem.
    f67 = open(arquivo67, encoding = 'utf-16')
    file_contents67 = f67.read()
    
    countAll1 = file_contents67.split('{"count":', 1)[1]
    countAll2 = countAll1.split(',"next"', 1)[0]
    
    index = 1
    
    while index <= int(countAll2):
        
        driver2.get('https://homolog.farmazon.com.br/api/products/?anvisa=&drug=&format=json&id__in=&image_url=&manufacturer=&max_price=&min_price=&name=&product_category='+categoryID+'&scientific_name=')
        
        with open('vitrine0.txt', 'w', encoding = 'utf-16') as f:
            f.write(driver2.page_source)
            
        arquivo68 = (r'C:\Users\Take4\Desktop\KPI\vitrine0.txt')
        
        #Arquivo TXT salvo é lido para ver quantas repeticoes do "create time" + ano + mes existem.
        f68 = open(arquivo68, encoding = 'utf-16')
        file_contents68 = f68.read()
    
        countProduct1 = file_contents68.split('json","id":', 1)[1]
        countProduct2 = countProduct1.split(',"name":', 1)[0]
    
        driver.get ('https://homolog.farmazon.com.br/admin/api/product/'+countProduct2)
        
        keyboard.press('ctrl')
        
        WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div[3]/div/form/div/fieldset/div[1]/div/div[1]/select/option['+categoryID+']'))).click()
        
        keyboard.release('ctrl')
        
        WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div/div[3]/div/form/div/div/input[3]'))).click()
        
        time.sleep(10)
        
        index += 1
    
    driver.close()
    driver2.close()
    
zerandoCategoria('646')
zerandoCategoria('647')
zerandoCategoria('648')


