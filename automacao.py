# Imports de bibliotecas 
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from openpyxl import load_workbook
from datetime import date
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
#-----------------------------------

openArquivo = load_workbook('formsAula.xlsx') #Leitura de arquivo para o openpyxl
aba = openArquivo.active #Apontamento de planilha ativa para o openpyxl

arquivo = pd.read_excel('formsAula.xlsx') #Leitura de arquivo para o pandas
print(arquivo) #Print de planilha no estado inicial
link = 'https://docs.google.com/forms/d/e/1FAIpQLSepRLCgDj1dZBE3FdwHZcS2qxzZQ2CLQGh-BY-SjwdcBgSaGA/viewform' #Link do forms ou site da ação

#Configuração para o chorme não abrir no front-end
opc = webdriver.ChromeOptions()
opc.add_argument("--headless")
#----------

chrm = webdriver.Chrome(ChromeDriverManager().install(),options=opc) #Abri o chorme

chrm.implicitly_wait(30) #Declara uma espera implícita, se caso ao elemento ou ação não for executada em 30 min daqui para frente, aborta o processo.
chrm.get(link) #Acessa o link definido
sleep(1)
for i, row in arquivo.iterrows(): #Loop para percorrer cada linha do excel e fazer a determinadas ações
    
    #Variáveis de uso na aplicação
    nome = row[0] #variável de nome
    nascimento = str(row[1]).split(' ')[0].split('-') #Transforma data em data arrey para a variável
    dia = nascimento[2] #Pega o dia do data arrey para a variável
    mes = nascimento[1] #Pega o mês do data arrey para a variável
    ano = nascimento[0] #Pega o ano do data arrey para a variável
    idade =  int(str(date.today()).split('-')[0]) - int(ano) #faz o calculo de idade ano atual menos variável ano
    dataNascimento = mes + dia + ano #Contatena as variáveis  dia, mês e ano para fora data de nascimento sem treços ou caracteres especiais
    Cargo = row[2] #variável de cargo de funcionário
    linha = i + 2 #variável de localização de linha para openpyxl
    #-------------------
    #Envia variáveis  para inputs definidos
    chrm.find_element(By.XPATH,'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(nome)
    chrm.find_element(By.XPATH,'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div/div[2]/div[1]/div/div[1]/input').send_keys(dataNascimento)
    chrm.find_element(By.XPATH,'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(idade)
    chrm.find_element(By.XPATH,'//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(Cargo)
    #----------------
    aba[f'E{linha}'] = 'OK' #Escreve ok na planilha com openpyxl para saber quais nomes foram escritos
    aba[f'D{linha}'] = chrm.find_element(By.XPATH,'//*[@id="mG61Hd"]/div[2]/div/div[1]/div/div[4]/div[2]').get_attribute('innerText')#Escreve atributo da web na planilha com openpyxl
    
    chrm.find_element(By.XPATH,'//*[@id="mG61Hd"]/div[2]/div/div[3]/div[1]/div[1]/div/span/span').click() #Envia o forms
    chrm.find_element(By.XPATH,'/html/body/div[1]/div[2]/div[1]/div/div[4]/a').click() #Volta para novo input de dados
    openArquivo.save('formsAula.xlsx') #Salva planilha 

    
chrm.quit() #Fecha o chorme

arquivo = pd.read_excel('formsAula.xlsx') #Leitura de arquivo para o pandas
print(arquivo) #Print de planilha no estado final

