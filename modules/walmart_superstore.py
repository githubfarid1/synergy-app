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
from openpyxl import Workbook, load_workbook
from urllib.parse import urlencode, urlparse
import time

def getConfig():
	file = open("setting.json", "r")
	config = json.load(file)
	return config


def browser_init():
    config = getConfig()
    warnings.filterwarnings("ignore", category=UserWarning)
    options = webdriver.ChromeOptions()
    # options = Options()
    # options.add_argument("--headless")
    options.add_argument("user-data-dir={}".format(config['chrome_user_data'])) 
    options.add_argument("profile-directory={}".format(config['chrome_profile']))
    options.add_argument('--no-sandbox')
    options.add_argument("--log-level=3")
    options.add_argument('--disable-blink-features=AutomationControlled')

    # options.add_argument("--window-size=1200, 900")
    options.add_argument('--start-maximized')
    options.add_argument("--disable-notifications")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    return webdriver.Chrome(service=Service(CM().install()), options=options)

driver = browser_init()
workbook = load_workbook(filename=r"C:/synergy-data-tester/Lookup Listing.xlsx", read_only=False, keep_vba=True, data_only=True)
# workbook = load_workbook(filename="/home/farid/dev/python/synergy-github/data/lookup/Lookup Listing.xlsx", read_only=False, keep_vba=True, data_only=True)
# worksheet = workbook[self.sheetname]
worksheet = workbook["Sheet1"]

for i in range(2, worksheet.max_row + 1):
    url = worksheet[f'A{i}'].value
    domain = urlparse(url).netloc
    if domain == 'www.walmart.com' or domain == 'www.walmart.ca':
        # driver.get(url)
        # time.sleep(3)
        input("wait")
        
         
