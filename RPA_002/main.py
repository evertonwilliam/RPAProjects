##########################################
# PRODUZIDO POR EVETON WILLIAM CONSTANTINO
# ROBÔ PARA DEMONSTRAR RPA DESAFIO 04
# DESAFIO 04
# DATA INICIAL: 18-11-2021
# DATA FINAL: 25/11/2021  
# CÓDIGO ABERTO PARA ESTUDOS
##########################################
# NECESSÁRIO OS SEGUINTES MÓDULOS PYTHON
# pip install Chatterbot
# py -m pip install --upgrade pip setuptools wheel
#   Certifique-se de instalar o pacote rarfile ( https://pypi.org/project/rarfile/ )
#   Certifique-se de instalar o winRar ( https://www.win-rar.com/ )
#   Adicione o caminho para winRar à variável de ambiente.
##########################################

## IMPORTAÇÃO DOS MÓDULOS PYTHON
import sqlite3
import time
import sys
import itertools
import threading
import os
import rarfile
import pyautogui
import logging
import logzero
from logzero import logger
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

# PASSO 01 - DEFINIÇÃO DE VARIÁVEIS
urlSite = 'https://mundorpa.com/index.html'
dirProj = 'C:\RPATemp\Python\Projetos_RPA_Python\Py_04'
dirLog  = '\log'
dirExe  = '\exe'
fileDw  = '\MundoRPA - Desafio App.rar'
fileEx  = '\MundoRPA - Desafio App.exe'
fileLogName = dirProj + dirLog + '\log.txt'

# Definição do processo de logs
logzero.logfile(fileLogName, maxBytes=1e6, backupCount=1, disableStderrLogger=False)
logger.setLevel(logging.DEBUG) 

# FUNÇÃO QUE CRIA PASTAS NO WINDOWS
def createFolders (folder):
    try:
        if os.path.exists(dirProj+folder) == False:
            logger.debug(':: Executar comando de criar diretórios')
            os.makedirs(dirProj+folder)
    except Exception as ex:
        print (ex)
        exit()
        
# FUNCÃO QUE VALIDA A EXISTENCIA DE ARQUIVOS EM PASTAS
def fileVerification(folderName, fileName):
    try:
        if os.path.exists(dirProj+folderName+fileName) == True:
            logger.debug(':: Executar o comando remove file')
            os.remove(dirProj+folderName+fileName)
    except Exception as ex:
        print(ex)
        exit()

# FUNCAO QUE INICIALIZA O NAVEGADOR
def initNavChrome(url):
    try:
        logger.debug(':: Iniciando o navegador e seus dados de navegação')
        options = webdriver.ChromeOptions()
        #prefs = {"download.default_directory" : dirProj+dirExe}
        prefs = {
                    "download.default_directory": dirProj+dirExe,
                    "download.prompt_for_download": False,
                    "download.directory_upgrade": True,
                    "safebrowsing.enabled": True
                }
        options.add_experimental_option("prefs",prefs)
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        options.add_argument("--disable-extensions")
        chrome = webdriver.Chrome(executable_path = 'chromedriver', options = options)
        chrome.get(url)
        return chrome
    except Exception as ex:
        logger.debug(':: Definir os dados do navegador chrome: '+ ex)
        print(ex)
        exit()

# FUNÇÃO QUE AGUARDA UM ELEMENTO EM TELA     
def waitElement (driver, type, object, time):
    try:
        logger.debug(':: Aguardar o elemento: ' + object)
        wait = WebDriverWait(driver, time)
        element = wait.until(EC.element_to_be_clickable((type, object)))
        return element
    except Exception as ex:
        logger.debug(':: Erro ao aguardar elemento: ' + ex)
        print(ex)
        exit()

# FUNÇÃO QUE EFETUA O DOWNLOAD DO ARQUIVO   
def getDownLoadedFileName(driver):
    driver.switch_to.window(driver.window_handles[-1])
    driver.get('chrome://downloads')
    logger.debug(':: Aguardar o download ocorrer')
    # CRIANDO A ANIMAÇÃO DE CARREGAMENTO
    for c in itertools.cycle(['|', '/', '-', '\\']):
        try:
            # CAPTURANDO O PERCENTUAL DE CARREGAMENTO
            downloadPercentage = driver.execute_script(
                "return document.querySelector('downloads-manager').shadowRoot.querySelector('#downloadsList downloads-item').shadowRoot.querySelector('#progress').value"
                )
                
            #CAPTURANDO O STATUS DO DOWNLOAD
            tagExecute = driver.execute_script(
                "return document.querySelector('body > downloads-manager').shadowRoot.querySelector('#frb0').shadowRoot.querySelector('#tag').getInnerHTML()"
               )
               
            # ANIMANDO O TERMINAL   
            sys.stdout.write('\rBaixando Arquivo ' + str(downloadPercentage) + '% ' + c + ' ')
            sys.stdout.flush()               
                
        except Exception as ex:
            return ex

        finally:
            # CASO FINALIZE O DOWNLOAD, CONTINUA O PROCESSO
            if downloadPercentage == 100:
                sys.stdout.write('\r')
                logger.debug(':: Confirmar o download do arquivo do site')
                time.sleep(1)
                pyautogui.hotkey('TAB')
                time.sleep(1)
                pyautogui.hotkey('TAB')
                time.sleep(1)
                pyautogui.hotkey('TAB')
                time.sleep(1)
                pyautogui.hotkey('TAB')
                time.sleep(1)
                pyautogui.hotkey('TAB')  
                time.sleep(1)
                pyautogui.hotkey('TAB') 
                time.sleep(1)
                pyautogui.hotkey('ENTER')
                time.sleep(2)
                pyautogui.hotkey('TAB')                
                time.sleep(1)
                pyautogui.hotkey('ENTER')
                time.sleep(5)

                driver.back()
                return 'Concluido'
            
            # CASO OCORRA ALGUM ERRO, RETORNA O ERRO
            if tagExecute == 'Falha - Erro na rede' or tagExecute == 'Cancelado':
                return tagExecute

# FUNCÃO QUE LÊ A TABELA DO SITE
def getTabelaWeb(driver, object):
    logger.debug(':: Ler os dados da tabela do site')
    itens = []
    for row in object.find_elements(By.XPATH, ".//tr"):
        itens.append([td.text for td in row.find_elements(By.XPATH, ".//td")])
    return itens

# EXECUTANDO O PROCESSO
if(__name__) == '__main__':
    
    # Cria os diretórios
    createFolders(dirLog)
    createFolders(dirExe)
    
    # verifica os downloads
    fileVerification(dirExe, fileDw)
    
    # Inicializa o navegador
    nav = initNavChrome(urlSite)
    
    # MAPEAMENTO
    logger.debug(':: Organizar as telas no desktop')
    pyautogui.hotkey('win', 'left')
    pyautogui.hotkey('esc')
 
    # LINK DESAFIOS
    linkDesafio = waitElement(nav, By.LINK_TEXT, 'Desafios', 40)
    linkDesafio.click()

    # LINK DE DOWNLOADS
    linkDownload = waitElement(nav, By.LINK_TEXT, 'Aplicação Desktop - Input', 30)
    linkDownload.click()

    # AGUARDANDO O DOWNLOAD E VERIFICANDO SE JÁ BAIXOU TUDO
    if getDownLoadedFileName(nav) != 'Concluido':
        logger.debug(':: Erro de Download')
        nav.close()
        exit()

    # LINK TABELA DE DADOS
    inkDesafio = waitElement(nav, By.LINK_TEXT, 'DESAFIO 4', 30)
    inkDesafio.click()
    
    # LER A TABELA
    webTable =  waitElement(nav, By.XPATH, "//*[@id='area-principal-desafio']/div/table/tbody", 30)
    dados = getTabelaWeb(nav, webTable)
    
    #FECHA NAVEGADOR
    logger.debug(':: Fechar o navegador')
    nav.close()

    # DESCOMPACTA O EXECUTÁVEL BAIXADO
    try:
        logger.debug(':: Descompactar o arquivo')
        r = rarfile.RarFile(dirProj+dirExe+fileDw)
        r.extractall(dirProj+dirExe)
        r.close()
    except Exception as ex:
        print(ex)
        exit()

    # EFETUA O CADASTRO
    try:
        logger.debug(':: Iniciar processo de cadastro no sistema')
        pyautogui.PAUSE = 0.1
        pyautogui.hotkey('winleft', 'r')
        pyautogui.hotkey('DEL')
        pyautogui.typewrite(dirProj+dirExe+fileEx)
        time.sleep(1)
        pyautogui.hotkey('ENTER')
        
        time.sleep(7)
        pyautogui.hotkey('TAB')    
    
        for i in range(len(dados)):
            logger.debug(':: Cadastrar o material: ' + dados[i][0])
            pyautogui.hotkey('TAB')
            pyautogui.hotkey('TAB')
            pyautogui.hotkey('TAB')    
            pyautogui.typewrite(dados[i][0])
            pyautogui.hotkey('TAB')
            pyautogui.typewrite(dados[i][1])
            pyautogui.hotkey('TAB')
            pyautogui.typewrite(dados[i][2])
            pyautogui.hotkey('TAB')
            pyautogui.hotkey('SPACE')
            pyautogui.hotkey('SPACE')
            
    except Exception as ex:
        logger.debug(':: Ocorreu erro no cadastro:' + ex)
        print(ex)
        exit()
    finally:
        logger.debug(':: ENCERRANDO O PROCESSO')
        pyautogui.hotkey('TAB')
        pyautogui.hotkey('TAB')
        pyautogui.hotkey('SPACE')
        pyautogui.hotkey('SPACE')        
        pyautogui.hotkey('ALT', 'F4')
        pyautogui.alert('Fim do Processo')
