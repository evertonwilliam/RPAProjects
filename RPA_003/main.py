from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

class chromeNav():

    

    def __init__(self):
        options = webdriver.ChromeOptions()
        prefs = {
                    "download.default_directory": "c:\Download",
                    "download.prompt_for_download": False,
                    "download.directory_upgrade": True,
                    "safebrowsing.enabled": True,
                    "excludeSwitches": ["enable-logging"]
                }
        options.add_experimental_option("prefs",prefs)
        options.add_argument("--disable-extensions")        

    def initNavChrome(self, url):
        try:
            chrome = webdriver.Chrome(executable_path = 'chromedriver', options = options)
            chrome.get(url)
            return chrome
        except Exception as ex:
            exit()
            
if __name__ == '__main__':
    app = chromeNav()
    app.initNavChrome("http://www.mercadolivre.com")