from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd

# Abri Navegador e entrar no Google
navegador = webdriver.Chrome()
td_cotas = ['Cotação Dolar', 'Cotação Euro']
# Cotação Dolar
navegador.get('https://www.google.com/')
pesquisa = navegador.find_element('xpath', '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input')
pesquisa.send_keys('Cotação Dolar')
pesquisa.send_keys(Keys.ENTER)
cota_dolar = navegador.find_element('xpath', '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]')
cota_dolar = cota_dolar.get_attribute('data-value') # Ler atributo de valor de cotação
print(cota_dolar)
# Cotação Euro
navegador.get('https://www.google.com/')
pesquisa = navegador.find_element('xpath', '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input')
pesquisa.send_keys('Cotação Euro')
pesquisa.send_keys(Keys.ENTER)
cota_euro = navegador.find_element('xpath', '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]')
cota_euro = cota_euro.get_attribute('data-value') # Ler atributo de valor de cotação
print(cota_euro)
# Cotação Ouro
navegador.get('https://www.melhorcambio.com/ouro-hoje')
cota_ouro = navegador.find_element('xpath', '//*[@id="comercial"]')
cota_ouro = cota_ouro.get_attribute('value') # Ler atributo de valor de cotação
cota_ouro = cota_ouro.replace(',', '.') # Substitui a virgula por ponto
print(cota_ouro)

navegador.quit()

# Atualizar a Base de Dados
tabela = pd.read_excel('Produtos.xlsx')
print(tabela)
print('='*100)

# Atualizando os preços
tabela.loc[tabela['Moeda'] == 'Dólar', 'Cotação'] = float(cota_dolar)
tabela.loc[tabela['Moeda'] == 'Euro', 'Cotação'] = float(cota_euro)
tabela.loc[tabela['Moeda'] == 'Ouro', 'Cotação'] = float(cota_ouro)
# Calculando os valores das colunas
tabela['Preço de Compra'] = tabela['Preço Original'] * tabela['Cotação']
tabela['Preço de Venda'] = tabela['Preço de Compra'] * tabela['Margem']
print(tabela)

tabela.to_excel('Produtos_novo.xlsx', index=False) # Exportando Tabela