##########################################
# PRODUZIDO POR EVETON WILLIAM CONSTANTINO
# ROBÔ PARA DEMONSTRAR ??????
# DESAFIO ????
# DATA INICIAL: 18-11-2021
# DATA FINAL:   
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
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By


#CONECTANDO EM UM BANCO DE DADOS SQLITE
#conn = sqlite3.connect('bot.db')

#Criando a tabela do banco de dados
#cur = conn.cursor()
#cur.execute('Create table if not exists bot(pergunta TEXT NULL, resposta TEXT NULL);') 
#cur.close()
#conn.close()


# PASSO 01 - DEFINIÇÃO DE VARIÁVEIS
urlSite = 'https://mundorpa.com/index.html'
dirProj = 'C:\RPATemp\Python\Projetos_RPA_Python\Py_04'
dirLog  = '\log'
dirExe  = '\exe'
fileDw  = '\MundoRPA - Desafio App.rar'
fileEx  = '\MundoRPA - Desafio App.exe'


# FUNÇÃO QUE CRIA PASTAS NO WINDOWS
def createFolders (folder):
    try:
        if os.path.exists(dirProj+folder) == False:
            os.makedirs(dirProj+folder)
    except Exception as ex:
        print (ex)
        exit()
        
# FUNCÃO QUE VALIDA A EXISTENCIA DE ARQUIVOS EM PASTAS
def fileVerification(folderName, fileName):
    try:
        if os.path.exists(dirProj+folderName+fileName) == True:
            os.remove(dirProj+folderName+fileName)
    except Exception as ex:
        print(ex)
        exit()

# FUNCAO QUE INICIALIZA O NAVEGADOR
def initNavChrome(url):
    try:
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
        print(ex)
        exit()

# FUNÇÃO QUE AGUARDA UM ELEMENTO EM TELA     
def waitElement (driver, type, object, time):
    try:
        wait = WebDriverWait(driver, time)
        element = wait.until(EC.element_to_be_clickable((type, object)))
        return element
    except Exception as ex:
        print(ex)
        exit()

# FUNÇÃO QUE EFETUA O DOWNLOAD DO ARQUIVO   
def getDownLoadedFileName(driver):
    driver.switch_to.window(driver.window_handles[-1])
    driver.get('chrome://downloads')

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
               
                driver.back()
                return 'Concluido'
            
            # CASO OCORRA ALGUM ERRO, RETORNA O ERRO
            if tagExecute == 'Falha - Erro na rede' or tagExecute == 'Cancelado':
                return tagExecute

# FUNCÃO QUE LÊ A TABELA DO SITE
def getTabelaWeb(driver, object):
    #table =  waitElement(driver, By.XPATH, object, 30)
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
    #fileVerification(dirExe, fileDw)
    
    # Inicializa o navegador
    nav = initNavChrome(urlSite)
    
    # MAPEAMENTO
 
    # LINK DESAFIOS
    linkDesafio = waitElement(nav, By.LINK_TEXT, 'Desafios', 40)
    linkDesafio.click()

    # LINK DE DOWNLOADS
    #linkDownload = waitElement(nav, By.LINK_TEXT, 'Aplicação Desktop - Input', 30)
    #linkDownload.click()

    # AGUARDANDO O DOWNLOAD E VERIFICANDO SE JÁ BAIXOU TUDO
    #if getDownLoadedFileName(nav) != 'Concluido':
        #print('Erro de Download')
        #nav.close()
        #exit()
    
    # LINK TABELA DE DADOS
    inkDesafio = waitElement(nav, By.LINK_TEXT, 'DESAFIO 4', 30)
    inkDesafio.click()
    
    # LER A TABELA
    webTable =  waitElement(nav, By.XPATH, "//*[@id='area-principal-desafio']/div/table/tbody", 30)
    dados = getTabelaWeb(nav, webTable)
    
    #FECHA NAVEGADOR
    nav.close()

    # DESCOMPACTA O EXECUTÁVEL BAIXADO
    try:
        r = rarfile.RarFile(dirProj+dirExe+fileDw)
        r.extractall(dirProj+dirExe)
        r.close()
    except Exception as ex:
        print(ex)
        exit()

    # EFETUA O CADASTRO
    try:
        pyautogui.PAUSE = 0.1
        pyautogui.hotkey('winleft', 'r')
        pyautogui.hotkey('DEL')
        pyautogui.typewrite(dirProj+dirExe+fileEx)
        time.sleep(1)
        pyautogui.hotkey('ENTER')
        
        time.sleep(7)
        pyautogui.hotkey('TAB')    
    
        for i in range(len(dados)):
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
        print(ex)
        exit()
    finally:
        pyautogui.hotkey('TAB')
        pyautogui.hotkey('TAB')
        pyautogui.hotkey('SPACE')
        pyautogui.hotkey('SPACE')        
        pyautogui.hotkey('ALT', 'F4')
        pyautogui.alert('Fim do Processo')
    
    
    
    
    
    