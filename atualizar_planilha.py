from selenium import webdriver
from selenium.webdriver import Keys
import pandas as pd
navegador = webdriver.Chrome()
navegador.get("https://www.google.com.br/")

#cotação dolar

navegador.find_element('xpath', '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("cotação do dolar")
navegador.find_element('xpath', '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
dolar = navegador.find_element('xpath','//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
#cotação euro 
navegador.get("https://www.google.com.br/")
navegador.find_element('xpath', '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("cotação do euro")
navegador.find_element('xpath', '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
euro = navegador.find_element('xpath','//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value')
#cotação ouro 
navegador.get("https://www.melhorcambio.com/ouro-hoje")
navegador.find_element('xpath', '/html/body/div[5]/div[1]/div/div/input[2]').send_keys("cotação do ouro")
ouro = navegador.find_element('xpath','/html/body/div[5]/div[1]/div/div/input[2]' ).get_attribute('value')
ouro = ouro.replace(",", ".")#troca a virgula por ponto

navegador.quit()#fecha navegador
#manipulando os valores na tabela
tabela = pd.read_excel("Produtos.xlsx")
#print(tabela)
tabela.loc[tabela['Moeda']=="Dólar", "Cotação"]=float(dolar) #substituindo os valores pelo novo buscado
tabela.loc[tabela['Moeda']=="Euro", "Cotação"]=float(euro)
tabela.loc[tabela['Moeda']=="Ouro", "Cotação"]=float(ouro)
#atualizar os valores de compra e venda
tabela["Preço de Compra"] = tabela["Preço Original"] * tabela["Cotação"]
tabela["Preço de Venda"] = tabela["Preço de Compra"] * tabela["Margem"]
tabela.to_excel("Produtos Novos.xlsx", index=False)#gera a nova tabela