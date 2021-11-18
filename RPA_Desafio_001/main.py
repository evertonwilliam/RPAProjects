##########################################
# PRODUZIDO POR EVETON WILLIAM CONSTANTINO
# ROBÔ PARA DEMONSTRAR O RPA EM PYTHON
# DESAFIO https://mundorpa.com/
# DATA INICIAL: 08-11-2021
# DATA FINAL:   11-11-2021
# CÓDIGO ABERTO PARA ESTUDOS
##########################################
# IMPORTANTES OS SEGUINTES MÓDULOS PYTHON
# PIP SELENIUM INSTALL
# PIP VALIDATE_EMAIL INSTALL
# PIP LOGZERO
# PIP PANDAS
##########################################

## IMPORTAÇÃO DOS MÓDULOS PYTHON
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from validate_email import validate_email
from logzero import logger
import logging
import logzero
import pandas
import time
import os

## PASSO 1 - DEFININDO AS VARIÁVEIS DE AMBIENTE
print(':: Definindo as variáveis')
urSite = 'https://mundorpa.com/index.html'
dirDownload = 'C:\RPATemp\Python\Desafios\MundoRpaDesafio_1'
dirLog = 'C:\RPATemp\Python\Desafios\MundoRpaDesafio_1\log'
fileName = dirDownload + '\Desafio1.xlsx'
fileLogName = dirLog + '\log.txt'

## Definição dos parametros do log
print(':: Definindo os parametros do log')
logzero.logfile(fileLogName, maxBytes=1e6, backupCount=1, disableStderrLogger=True)
logger.setLevel(logging.DEBUG) 

## PASSO 2 - CRIANDO DIRETÓRIOS CASO NÃO EXISTA
try:
    ## Criar os diretórios
    print(':: Verificando se os diretórios do processo existem')
    if os.path.exists(dirDownload) == False:
        print(':: Criar o diretório do Excel')
        logger.debug(':: Criar o diretório do Excel')
        os.makedirs(dirDownload)

    if os.path.exists(dirLog) == False:
        print(':: Criando o diretório do Log')
        logger.debug(':: Criando o diretório do Log')    
        os.makedirs(dirLog)
        
except Exception as ex:
    print(':: Erro ao criar diretórios')
    logger.exception(':: Erro ao criar diretórios' + ex)
    exit()
    

## PASSO 3 - VERIFICANDO SE O ARQUIVO JÁ FOI BAIXADO
try:
    ## Preparando o ambiente
    print(':: Verificando se o arquivo existe')
    if os.path.exists(fileName):
        print(':: Removendo o arquivo')
        logger.debug(':: Criar o diretório do Log')
        os.remove(fileName)
except Exception as ex:
    print(':: Erro ao remover o arquivo existente')
    logger.exception(':: Erro ao remover o arquivo existente ' + ex)
    exit()
    

## PASSO 4 - INICIANDO O NAVEGADOR
try:
    print(':: Iniciando o navegador')
    logger.debug(':: Iniciando o navegador')
    options = webdriver.ChromeOptions()
    prefs = {"download.default_directory" : dirDownload}
    options.add_experimental_option("prefs",prefs)
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.add_argument("--disable-extensions")
    chrome = webdriver.Chrome(executable_path = 'chromedriver', options = options)
    chrome.get(urSite)
except Exception as ex:
    print(':: Erro ao inicializar o navegador')
    logger.exception(':: Erro ao inicializar o navegador ' + ex)
    exit()  

    
## PASSO 5 - NAVEGANDO NOS LINKS DO SITE
try:
    print(':: Clicando no link Desafios')
    logger.debug(':: Clicando no link Desafios')
    linkDesafios = chrome.find_element(By.LINK_TEXT, 'Desafios')
    linkDesafios.click()
    time.sleep(2)
except Exception as ex:
    print(':: Link Desafios não foi identificado')
    logger.exception(':: Link Desafios não foi identificado ' + ex)
    exit()

## Passo 6 - Efetuando o download do arquivo
try:
    print(':: Efetuando o download do arquivo')
    logger.debug(':: Efetuando o download do arquivo')
    linkDownload = chrome.find_element(By.CSS_SELECTOR, '#area-principal > p:nth-child(12) > a')
    linkDownload.click()
    time.sleep(1)
except Exception as ex:
    print(':: Ocorreu um erro durante o download')
    logger.exception(':: Ocorreu um erro durante o download ' + ex)
    exit()

## Passo 7 - Navegando até o formulário de cadastro
try:
    print(':: Clicando no link para acesso ao formulário de Cadastro')
    logger.debug(':: Clicando no link para acesso ao formulário de Cadastro')
    linkDesafio1 = chrome.find_element(By.LINK_TEXT, 'DESAFIO 1')
    linkDesafio1.click()
    time.sleep(1)
except:
    print(':: O link do formulário não foi identificado')
    logger.exception(':: O link do formulário não foi identificado' + ex)
    exit()
    
    
## Passo 8 - Abrindo e lendo o arquivo Excel
try:
    print(':: Abrindo e lendo o arquivo Excel')
    logger.debug(':: Abrindo e lendo o arquivo Excel')
    dados = pandas.read_excel(fileName, sheet_name='DesafioInput1')
except:
    print(':: O arquivo Excel não pode ser lido')
    logger.exception(':: O arquivo Excel não pode ser lido' + ex)
    exit()
    
## Passo 9 - Cadastrando os dados dentro do site
try:
    for row in range(len(dados)):
        
        print(':: Cadastrando a pessoa: ', dados['Nome Completo'][row])
        logger.debug(':: Cadastrando a pessoa ' + dados['Nome Completo'][row])
        
        print(':: Validando o e-mail: ', dados['Email'][row])
        valid = validate_email(dados['Email'][row])
        
        if valid:
            nome = chrome.find_element(By.ID, 'nomeCompleto')
            nome.send_keys(dados['Nome Completo'][row])

            pNome = chrome.find_element(By.ID, 'pNome')
            pNome.send_keys(dados['Primeiro Nome'][row])

            uNome = chrome.find_element(By.ID, 'Unome')
            uNome.send_keys(dados['Ultimo Nome'][row])
            
            email = chrome.find_element(By.ID, 'email')
            email.send_keys(dados['Email'][row])

            dtNascimento = chrome.find_element(By.ID, 'dtNascimento')
            dtNascimento.send_keys(dados['Data de Nascimento'][row])

            telefone = chrome.find_element(By.ID, 'telefone')
            telefone.send_keys(dados['Telefone'][row])

            btEnviar = chrome.find_element(By.XPATH,'//*[@id="formulario1"]/input[7]')
            btEnviar.click()
            
            time.sleep(1)
        else:
            print('::::- Email:', dados['Email'][row], 'é Invalido')
            print('::::- Pessoa:', dados['Nome Completo'][row], 'Não Cadastrada')
            logger.warning('::::- Email: ', dados['Email'][row], ' é Invalido ')
except:
    print(':: Ocorreu um erro no processo de cadastro do formulário')
    logger.exception(':: Ocorreu um erro no processo de cadastro do formulário' + ex)
    exit()


## Passo 10 - Encerrando o processo
print(':: Fim do processo')
chrome.close()
