from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager as CM
from selenium.webdriver.common.action_chains import ActionChains
import json
import warnings
from urllib.parse import urlencode, urlparse
import time
from random import randint
import pyautogui as pg
import undetected_chromedriver as uc 
import os
import shutil
import xlwings as xw

def getConfig():
	file = open("setting.json", "r")
	config = json.load(file)
	return config


def browser_init():
    config = getConfig()
    warnings.filterwarnings("ignore", category=UserWarning)
    options = webdriver.ChromeOptions()
    
    options.add_argument("user-data-dir={}".format("C:\\Users\\User\\AppData\\Local\\Google\\Chrome\\User Data2")) 
    options.add_argument("profile-directory={}".format("Default"))

    options.add_argument('--no-sandbox')
    options.add_argument("--log-level=3")
    # options.add_argument("--window-size=1200, 900")
    options.add_argument('--start-maximized')
    options.add_argument("--disable-notifications")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument("--disable-blink-features=AutomationControlled")
    # options.add_experimental_option( "prefs",{'profile.managed_default_content_settings.javascript': 1})
    driver = webdriver.Chrome(service=Service(CM().install()), options=options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})") 
    driver.execute_cdp_cmd("Network.setCacheDisabled", {"cacheDisabled":True})
    return driver


filename = r"C:/synergy-data-tester/Lookup Listing.xlsx"
sheetname = "Sheet1"
xlbook = xw.Book(filename)
xlsheet = xlbook.sheets[sheetname]

user_data = r"C:/Users/User/AppData/Local/Google/Chrome/User Data2"

def get_urls(domainwl=[]):
    urlList = []
    print(domainwl)
    maxrow = xlsheet.range('A' + str(xlsheet.cells.last_cell.row)).end('up').row
    for i in range(2, maxrow + 2):
        url = xlsheet[f'A{i}'].value
        domain = urlparse(url).netloc
        # if domain in 'www.walmart.com' or domain == 'www.walmart.ca':
        if domain in domainwl:
            tpl = (url, i)
            urlList.append(tpl)
    return urlList

def walmart_scraper():
    urlList = get_urls(domainwl=['www.walmart.com','www.walmart.ca'])
    i = 0
    maxrec = len(urlList)
    driver = browser_init()
    while True:
        if i == maxrec:
            break
        url = urlList[i][0]
        rownum = urlList[i][1]
        print(url, end=" ", flush=True)
        driver.get(url)
        try:
            driver.find_element(By.CSS_SELECTOR, "div#topmessage").text
            print("Failed")
            del driver
            waiting = 120
            print(f'The script was detected as bot, please wait for {waiting} seconds', end=" ", flush=True)
            time.sleep(waiting)
            isExist = os.path.exists(user_data)
            print(isExist)
            if isExist:
                shutil.rmtree(user_data)
            print('OK')
            driver = browser_init()
            continue
        except:
            
            print('OK')
            pass

        try:
            title = driver.find_element(By.CSS_SELECTOR, "h1[data-automation='product-title']").text
        except:
                title = ''
        try:
            price = driver.find_element(By.CSS_SELECTOR, "span[data-automation='buybox-price']").text
        except:
            price = ''
        try:
            sale = driver.find_element(By.CSS_SELECTOR, "div[data-automation='mix-match-badge'] span").text
        except:
            sale = ''
        
        print(title, price, sale)
        
        xlsheet[f'B{rownum}'].value = price
        xlsheet[f'C{rownum}'].value = sale

        i += 1     

    xlbook.save(filename)


def superstore_scraper():
    urlList = get_urls(domainwl=['www.realcanadiansuperstore.ca'])
    i = 0
    maxrec = len(urlList)
    driver = browser_init()
    while True:
        if i == maxrec:
            break
        url = urlList[i][0]
        rownum = urlList[i][1]
        print(url, end=" ", flush=True)
        driver.get(url)
        # try:
        #     driver.find_element(By.CSS_SELECTOR, "div#topmessage").text
        #     print("Failed")
        #     del driver
        #     waiting = 120
        #     print(f'The script was detected as bot, please wait for {waiting} seconds', end=" ", flush=True)
        #     time.sleep(waiting)
        #     isExist = os.path.exists(user_data)
        #     print(isExist)
        #     if isExist:
        #         shutil.rmtree(user_data)
        #     print('OK')
        #     driver = browser_init()
        #     continue
        #     raise
        # except:
        #     print('OK')
        #     pass
        # print('OK')
        # try:
        #     title = driver.find_element(By.CSS_SELECTOR, "h1[class='product-name__item product-name__item--name']").text
        # except:
        #     title = 'xx'

        # try:
        #     price = driver.find_element(By.CSS_SELECTOR, "span[class='price__value selling-price-list__item__price selling-price-list__item__price--now-price__value']").text
        # except:
        #     price = ''
        # try:
        #     # sale = driver.find_element(By.CSS_SELECTOR, "div[data-automation='mix-match-badge'] span").text
        #     raise
        # except:
        #     sale = ''
        
        # print(title, price, sale)
        
        # xlsheet[f'B{rownum}'].value = price
        # xlsheet[f'C{rownum}'].value = sale
        input('wa')
        i += 1
        time.sleep(5)

    # xlbook.save(filename)

if __name__ == '__main__':
    superstore_scraper()


# span price__value selling-price-list__item__price selling-price-list__item__price--now-price__value
# h1 product-name__item product-name__item--name