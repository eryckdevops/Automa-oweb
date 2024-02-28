from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
import pandas as pd

s = Service(r'')
driver = webdriver.Chrome(service=s)

file_path = r''
df = pd.read_excel(file_path)
organizacoes = df['Organização'].dropna()

driver.get("")

for organizacao in organizacoes:
    try:
        campo_nome = driver.find_element(By.NAME, 'Name')
        campo_nome.clear()
        campo_nome.send_keys(organizacao)
        
        driver.find_element(By.ID, 'btn-submit').click()
        
        time.sleep(5)
        
        driver.get("")
        
    except Exception as e:
        print(f"Erro ao preencher o formulário para {organizacao}: {e}")
        driver.close() 
        driver.switch_to.window(driver.window_handles[0]) 
        continue

driver.quit()
