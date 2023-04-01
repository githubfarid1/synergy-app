# import settings
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager as CM
import time
from openpyxl import Workbook, load_workbook
# If you need to get the column letter, also import this
from openpyxl.utils import get_column_letter
# from bs4 import BeautifulSoup
import warnings
import argparse
import os

def parse(fileinput, chrome_data):
    warnings.filterwarnings("ignore", category=UserWarning) 

    trackupdate_source = fileinput
    # url = "https://sellercentral.amazon.com/orders-v3/mfn/unshipped/ref=xx_orders_cont_kpiToolbar?_encoding=UTF8?communicationDeliveryId=732ae59b-0a70-42ac-8c7f-fc3b877b81aa&mons_sel_dir_mcid=amzn1.merchant.d.ABVP3LFM3UIGGSE7SBSDGWYKX47Q&mons_sel_mkid=ATVPDKIKX0DER&shipByDate=all&page=1"

    options = webdriver.ChromeOptions()
    # options.add_argument("--headless")
    # options.add_experimental_option('debuggerAddress', 'localhost:9251')
    options.add_argument("user-data-dir={}".format(chrome_data)) #Path to your chrome profile
    options.add_argument('--no-sandbox')
    options.add_argument("--log-level=3")
    options.add_argument("--window-size=800,600")
    # options.add_argument("user-agent=" + ua.random )
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    driver = webdriver.Chrome(service=Service(CM().install()), options=options)
    os.system('cls')
    print('File Selected:', fileinput)

    wb = load_workbook(filename=trackupdate_source, read_only=False, keep_vba=True, data_only=True)
    ws = wb['dyk_manifest_template']
    # Use the active cell when the file was loaded
    ws = wb.active
    first = True
    for i in range(2, ws.max_row + 1):
        if ws['R{}'.format(i)].value == None:
            break
        order_id = ws['R{}'.format(i)].value
        # ws['I{}'.format(i)].value = ws['B{}'.format(i)].value + ' ' + ws['C{}'.format(i)].value + ' ' + ws['D{}'.format(i)].value
        url = 'https://sellercentral.amazon.com/orders-v3/order/{}'.format(order_id) # 111-9589748-6199459
        driver.get(url)
        time.sleep(2)
        print('order ID',order_id)
        try:
            
            # WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH , '//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[8]/div/div/div[1]/div[2]/div/div[2]/div[2]/div[2]/div/span/a')))
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "a[data-test-id='tracking-id-value']")))
        
            
            # tracking_id = driver.find_element(By.XPATH,'//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[8]/div/div/div[1]/div[2]/div/div[2]/div[2]/div[2]/div/span/a').text
            tracking_id = driver.find_element(By.CSS_SELECTOR, "a[data-test-id='tracking-id-value']").text

                                                
            try:
                weight = driver.find_element(By.XPATH,'//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[8]/div/div/div[1]/div[3]/div/div/div[2]/div[2]/div[3]/div/div[2]').text
            except:
                weight = driver.find_element(By.XPATH,'//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/div[1]/div[1]/div[3]/div/div/div[2]/div[2]/div[3]/div/div[2]').text

            # SHIPPING COST
            try:
                cost = driver.find_element(By.XPATH,'//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[8]/div/div/div[1]/div[3]/div/div/div[2]/div[3]/div/div[2]/div[2]/span').text.replace('$','')
                
            except:
                # 114-5921481-0720211
                try:
                    cost = driver.find_element(By.XPATH,'//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/table/tbody/tr/td[7]/div/table[1]/tbody/div[3]/div[2]/span').text.replace('$','')
                except:
                    try:
                        cost = driver.find_element(By.XPATH,'//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/div/div[1]/div[3]/div/div/div[2]/div[3]/div/div[2]/div[2]/span').text.replace('$','')
                    except:
                        cost = ''
            
            # SHIPPING SERVICE
            try:
                service = driver.find_element(By.XPATH,'//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[8]/div/div/div[1]/div[2]/div/div[3]/div/div[2]').text
            except:
                try:
                    service = driver.find_element(By.XPATH,'//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/div/div[1]/div[2]/div/div[3]/div/div[2]').text
                except:
                    service = ''
            
            print(tracking_id,weight, cost, service)
            ws['M{}'.format(i)].value = tracking_id
            ws['N{}'.format(i)].value = weight
            ws['O{}'.format(i)].value = cost
            ws['P{}'.format(i)].value = service




            # print(timetr, loctr, eventtr)

        except:
            print('failed')
            ws['M{}'.format(i)].value = ''
            ws['N{}'.format(i)].value = ''
            ws['O{}'.format(i)].value = ''
            ws['P{}'.format(i)].value = ''

        try:
            try:
                trackbutton = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[7]/div/div/div[1]/div[2]/div/div[2]/div[2]/div[2]/div/span/a/i'))).click()
            except:
                trackbutton = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="MYO-app"]/div/div[1]/div[1]/div/div[8]/div/div/div[1]/div[2]/div/div[2]/div[2]/div[2]/div/span/a/i'))).click()
            # input('wait')
            
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@id="a-popover-content-1"]/div/table/tr[1]')))
            time.sleep(1.5)
            tracktable = driver.find_element(By.XPATH, '//*[@id="a-popover-content-1"]/div/table/tr[2]').find_elements(By.CSS_SELECTOR, 'td')
            timetr = tracktable[0].text
            loctr = tracktable[1].text
            eventtr = tracktable[2].text
            print(timetr)
            ws['Z{}'.format(i)].value = timetr
            ws['AA{}'.format(i)].value = loctr
            ws['AB{}'.format(i)].value = eventtr
        except:
            ws['Z{}'.format(i)].value = ''
            ws['AA{}'.format(i)].value = ''
            ws['AB{}'.format(i)].value = ''

        # input('wait')
        
        # html = driver.execute_script("return document.getElementsByTagName('html')[0].innerHTML")
        # soup = BeautifulSoup(html,"html.parser")
        # tracktable = soup.find('div', class_='a-popover-content')
        # if tracktable != None:
        #     print(tracktable.find('table', class_='a-normal').text)
        # input('d')
    wb.save(trackupdate_source)

    input('Process Finished...')

def main():
    parser = argparse.ArgumentParser(description="Tracking Update")
    parser.add_argument('-i', '--input', type=str,help="File Input")
    parser.add_argument('-d', '--data', type=str,help="Chrome User Data Directory")
    args = parser.parse_args()
    if not (args.input[-5:] == '.xlsx' or args.input[-5:] == '.xlsm'):
        print('File input have to XLSM or XLSX file')
        exit()
    isExist = os.path.exists(args.data)
    if isExist == False :
        print('Please check Chrome User Data Directory')
        exit()
    isExist = os.path.exists(args.input)
    if isExist == False :
        print('Please check XLSX file')
        exit()
    
    parse(args.input, args.data)
    
if __name__ == '__main__':
    main()

