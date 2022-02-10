#bibliotecas necessarias
import pandas as pd                             
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import urllib
#importa arquivo com o numero,
df = pd.read_excel('contatos.xlsx')
#iniciar navegador 
navegador = webdriver.Chrome()
#abrir pagina 
navegador.get("https://web.whatsapp.com/")
#esperar fazer login  
while len(navegador.find_elements_by_id("side")) <1 :
    time.sleep(1)
#login feito entao: 
for i, horario in enumerate(df["horario"]): #i = indice do cliente
    #armazena dados da planilha nas variaveis 
    pessoas = df.loc[i, "cliente"]
    numero = df.loc[i, "numero"]
    loja = df.loc[i,"loja"]
    horario = df.loc[i,"horario"]
    #converter texto para formato URL
    texto = urllib.parse.quote(f"Ola {pessoas}, Apoio Security informa que o sistema de alarme localizado no(a) {loja} está desativado. estava agendado para ser acionado as {horario}, caso já tenhaa ativado desconsidere essa mensagem. ")
    #cria o link para enviar as mensagens 
    link = f"https://web.whatsapp.com/send?phone={numero}&text={texto}"
    #abre novo link para enviar as mensagens 
    navegador.get(link)
    time.sleep(5)
    #click no botao de envio 
    navegador.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span').click()
    time.sleep(10)
    