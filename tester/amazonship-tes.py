from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager as CM
from selenium.webdriver.support.select import Select
import sys

datalist = [{'begin': 3, 'end': 25, 'submitter': 'DMT Distributors (102 Central Avenue)', 'address': 'SCK4 (102 Central Ave Ste 6540)', 'name': 'Shipment 1', 'boxcount': 3, 'weightboxes': [46, 46, 15], 'dimensionboxes': ['16x16x16', '16x16x16', '14x14x14'], 'items': [{'id': 'U17', 'name': 'Mars Maltesers Celebration Size 800g/1.7lbs. Bag {Imported from Canada}', 'total': 58, 'expiry': '2023-02-12 00:00:00', 'boxes': [29, 29, 0]}, {'id': 'U171', 'name': 'Real Jubes ORIGINAL 818g 4.16oz', 'total': 10, 'expiry': '2023-04-12 00:00:00', 'boxes': [0, 5, 5]}]}, {'begin': 26, 'end': 48, 'submitter': 'DMT Distributors (102 Central Avenue)', 'address': 'FTW1 (4160 Feathergrass Ln Prosper,  TX  75078)', 'name': 'Shipment 2', 'boxcount': 1, 'weightboxes': [35], 'dimensionboxes': ['18x18x18'], 'items': [{'id': 'B87', 'name': 'Nestle Coffin Crisp Coffee Crisp 30x12g Snack Size Bars - Imported From Canada', 'total': 40, 'expiry': '2022-02-19 00:00:00', 'boxes': [40]}]}, {'begin': 49, 'end': 71, 'submitter': 'DMT Distributors (102 Central Avenue)', 'address': 'MQJ1 (2005 Kansas Avenue, Flint, MI, 48502)', 'name': 'Shipment 3', 'boxcount': 1, 'weightboxes': [35], 'dimensionboxes': ['18x18x18'], 'items': [{'id': 'b87', 'name': 'Nestle Coffin Crisp Coffee Crisp 30x12g Snack Size Bars - Imported From Canada', 'total': 40, 'expiry': '2022-02-19 00:00:00', 'boxes': [40]}]}, {'begin': 72, 'end': 94, 'submitter': 'DMT Distributors (102 Central Avenue)', 'address': 'LAS1 (4020 west cambridge avenue Phoenix,  AZ  85009)', 'name': 'Shipment 4', 'boxcount': 1, 'weightboxes': [35], 'dimensionboxes': ['18x18x18'], 'items': [{'id': 'B87', 'name': 'Nestle Coffin Crisp Coffee Crisp 30x12g Snack Size Bars - Imported From Canada', 'total': 32, 'expiry': '2022-02-19 00:00:00', 'boxes': [40]}, {'id': 'U966', 'name': 'Nestle After Eight Collection Dark And White Mint Chocolates 150g', 'total': 19, 'expiry': '2022-05-07 00:00:00', 'boxes': [19]}]}, {'begin': 95, 'end': 117, 'submitter': 'DMT Distributors (102 Central Avenue)', 'address': 'CLT2 (746 NW 7th Ave Florida City,  FL  33034)', 'name': 'Shipment 5', 'boxcount': 1, 'weightboxes': [35], 'dimensionboxes': ['18x18x18'], 'items': [{'id': 'B87', 'name': 'Nestle Coffin Crisp Coffee Crisp 30x12g Snack Size Bars - Imported From Canada', 'total': 40, 'expiry': '2022-02-19 00:00:00', 'boxes': [40]}]}, {'begin': 118, 'end': 140, 'submitter': 'DMT Distributors (102 Central Avenue)', 'address': 'ABE8 (118 Whitmore Ave Wayne,  NJ  07470)', 'name': 'Shipment 6', 'boxcount': 1, 'weightboxes': [31], 'dimensionboxes': ['18x18x18'], 'items': [{'id': 'U846', 'name': 'NESTLÉ MINIS Assorted Bars - KITKAT, Coffee Crisp, AERO, Smarties - 303g (Pack of 30 Mini Bars)', 'total': 40, 'expiry': '2023-06-06 00:00:00', 'boxes': [40]}]}, {'begin': 141, 'end': 163, 'submitter': 'DMT Distributors (102 Central Avenue)', 'address': 'MEM1 (248 Cargile Ln Nashville,  TN  37205)', 'name': 'Shipment 7', 'boxcount': 2, 'weightboxes': [31, 31], 'dimensionboxes': ['18x18x18', '18x18x18'], 'items': [{'id': 'U846', 'name': 'NESTLÉ MINIS Assorted Bars - KITKAT, Coffee Crisp, AERO, Smarties - 303g (Pack of 30 Mini Bars)', 'total': 80, 'expiry': '2023-06-06 00:00:00', 'boxes': [40, 40]}]}, {'begin': 164, 'end': 186, 'submitter': 'Douglas Tsang (102 Central Avenue)', 'address': 'SCK4 (102 Central Ave Ste 6540)', 'name': 'Shipment 8', 'boxcount': 1, 'weightboxes': [46], 'dimensionboxes': ['16x16x16'], 'items': [{'id': 'U1344', 'name': 'Neon Tropical Candy Powder Filled Straws', 'total': 37, 'expiry': '2024-12-31 00:00:00', 'boxes': [37]}, {'id': 'U1498', 'name': 'Neon Lazers Candy Powder Filled Straws', 'total': 2, 'expiry': '2024-12-31 00:00:00', 'boxes': [2]}]}, {'begin': 187, 'end': 209, 'submitter': 'Douglas Tsang (102 Central Avenue)', 'address': 'FTW1 (4160 Feathergrass Ln Prosper,  TX  75078)', 'name': 'Shipment 9', 'boxcount': 1, 'weightboxes': [46], 'dimensionboxes': ['16x16x16'], 'items': [{'id': 'U1159', 'name': 'Neon Candy Powder Filled Straws, 120 Count (Sour)', 'total': 39, 'expiry': '2024-12-31 00:00:00', 'boxes': [39]}]}, {'begin': 210, 'end': 232, 'submitter': 'Douglas Tsang (102 Central Avenue)', 'address': 'MQJ1 (2005 Kansas Avenue, Flint, MI, 48502)', 'name': 'Shipment 10', 'boxcount': 1, 'weightboxes': [46], 'dimensionboxes': ['16x16x16'], 'items': [{'id': 'U1159', 'name': 'Neon Candy Powder Filled Straws, 120 Count (Sour)', 'total': 39, 'expiry': '2024-12-31 00:00:00', 'boxes': [39]}]}, {'begin': 233, 'end': 255, 'submitter': 'Douglas Tsang (102 Central Avenue)', 'address': 'LAS1 (4020 west cambridge avenue Phoenix,  AZ  85009)', 'name': 'Shipment 11', 'boxcount': 1, 'weightboxes': [46], 'dimensionboxes': ['18x18x18'], 'items': [{'id': 'U952', 'name': 'Nestle Turtles Original; Limited Edition; 333g/11.7oz., Tin (Imported from Canada)', 'total': 32, 'expiry': '2023-05-30 00:00:00', 'boxes': [32]}]}, {'begin': 256, 'end': 278, 'submitter': 'Douglas Tsang (102 Central Avenue)', 'address': 'CLT2 (746 NW 7th Ave Florida City,  FL  33034)', 'name': 'Shipment 12', 'boxcount': 1, 'weightboxes': [46], 'dimensionboxes': ['18x18x18'], 'items': [{'id': 'U952', 'name': 'Nestle Turtles Original; Limited Edition; 333g/11.7oz., Tin (Imported from Canada)', 'total': 32, 'expiry': '2023-05-30 00:00:00', 'boxes': [32]}]}, {'begin': 279, 'end': 301, 'submitter': 'Douglas Tsang (102 Central Avenue)', 'address': 'ABE8 (118 Whitmore Ave Wayne,  NJ  07470)', 'name': 'Shipment 13', 'boxcount': 1, 'weightboxes': [31], 'dimensionboxes': ['20x20x20'], 'items': [{'id': 'u841', 'name': 'Hersheys Cookies n Creme Advent Calendar 212g/ 7.4oz', 'total': 28, 'expiry': '2023-05-28 00:00:00', 'boxes': [28]}]}, {'begin': 302, 'end': 324, 'submitter': 'Douglas Tsang (102 Central Avenue)', 'address': 'MEM1 (248 Cargile Ln Nashville,  TN  37205)', 'name': 'Shipment 14', 'boxcount': 1, 'weightboxes': [31], 'dimensionboxes': ['20x20x20'], 'items': [{'id': 'u841', 'name': 'Hersheys Cookies n Creme Advent Calendar 212g/ 7.4oz', 'total': 28, 'expiry': '2023-05-28 00:00:00', 'boxes': [28]}]}, {'begin': 325, 'end': 347, 'submitter': 'DRCR (102 Central Avenue)', 'address': 'SCK4 (102 Central Ave Ste 6540)', 'name': 'Shipment 15', 'boxcount': 1, 'weightboxes': [26], 'dimensionboxes': ['20x20x20'], 'items': [{'id': 'u1501', 'name': 'KITKAT Advent Calendar Chocolate Candy, 270g/9.5oz {Imported from Canada}', 'total': 2, 'expiry': '2023-05-01 00:00:00', 'boxes': [2]}, {'id': 'U841', 'name': 'Hersheys Cookies n Creme Advent Calendar 212g/ 7.4oz', 'total': 26, 'expiry': '2023-05-28 00:00:00', 'boxes': [26]}]}, {'begin': 348, 'end': 370, 'submitter': 'DRCR (102 Central Avenue)', 'address': 'FTW1 (4160 Feathergrass Ln Prosper,  TX  75078)', 'name': 'Shipment 16', 'boxcount': 1, 'weightboxes': [24], 'dimensionboxes': ['20x20x20'], 'items': [{'id': 'U759', 'name': 'Nestle After Eight Advent Calendar Santa Dark Chocolate, 199g/7oz {Imported from Canada}', 'total': 27, 'expiry': '2023-03-13 00:00:00', 'boxes': [27]}]}, {'begin': 371, 'end': 393, 'submitter': 'DRCR (102 Central Avenue)', 'address': 'MQJ1 (2005 Kansas Avenue, Flint, MI, 48502)', 'name': 'Shipment 17', 'boxcount': 1, 'weightboxes': [24], 'dimensionboxes': ['20x20x20'], 'items': [{'id': 'U759', 'name': 'Nestle After Eight Advent Calendar Santa Dark Chocolate, 199g/7oz {Imported from Canada}', 'total': 27, 'expiry': '2023-03-13 00:00:00', 'boxes': [27]}]}, {'begin': 394, 'end': 416, 'submitter': 'DRCR (102 Central Avenue)', 'address': 'LAS1 (4020 west cambridge avenue Phoenix,  AZ  85009)', 'name': 'Shipment 18', 'boxcount': 1, 'weightboxes': [33], 'dimensionboxes': ['20x20x20'], 'items': [{'id': 'u1501', 'name': 'KITKAT Advent Calendar Chocolate Candy, 270g/9.5oz {Imported from Canada}', 'total': 28, 'expiry': '2023-05-01 00:00:00', 'boxes': [28]}]}, {'begin': 417, 'end': 439, 'submitter': 'DRCR (102 Central Avenue)', 'address': 'CLT2 (746 NW 7th Ave Florida City,  FL  33034)', 'name': 'Shipment 19', 'boxcount': 1, 'weightboxes': [42], 'dimensionboxes': ['20x20x20'], 'items': [{'id': 'u1501', 'name': 'KITKAT Advent Calendar Chocolate Candy, 270g/9.5oz {Imported from Canada}', 'total': 23, 'expiry': '2023-05-01 00:00:00', 'boxes': [23]}, {'id': 'u17', 'name': 'Mars Maltesers Celebration Size 800g/1.7lbs. Bag {Imported from Canada}', 'total': 1, 'expiry': '2023-02-12 00:00:00', 'boxes': [1]}, {'id': 'u952', 'name': 'Nestle Turtles Original; Limited Edition; 333g/11.7oz., Tin (Imported from Canada)', 'total': 8, 'expiry': '2023-05-30 00:00:00', 'boxes': [8]}]}, {'begin': 440, 'end': 462, 'submitter': 'DRCR (102 Central Avenue)', 'address': 'ABE8 (118 Whitmore Ave Wayne,  NJ  07470)', 'name': 'Shipment 20', 'boxcount': 1, 'weightboxes': [26], 'dimensionboxes': ['20x20x20'], 'items': [{'id': 'U1501', 'name': 'KITKAT Advent Calendar Chocolate Candy, 270g/9.5oz {Imported from Canada}', 'total': 6, 'expiry': '2023-05-01 00:00:00', 'boxes': [6]}, {'id': 'U1502', 'name': 'Nestle Kit Kat Christmas Holiday Chocolate Advent Calendar, 208g/7.3 oz. Box {Imported from Canada}', 'total': 20, 'expiry': '2023-05-23 00:00:00', 'boxes': [20]}]}, {'begin': 463, 'end': 485, 'submitter': 'DRCR (102 Central Avenue)', 'address': 'MEM1 (248 Cargile Ln Nashville,  TN  37205)', 'name': 'Shipment 21', 'boxcount': 1, 'weightboxes': [46], 'dimensionboxes': ['16x16x16'], 'items': [{'id': 'U1159', 'name': 'Neon Candy Powder Filled Straws, 120 Count (Sour)', 'total': 22, 'expiry': '2024-12-31 00:00:00', 'boxes': [22]}, {'id': 'U1344', 'name': 'Neon Tropical Candy Powder Filled Straws', 'total': 16, 'expiry': '2024-12-31 00:00:00', 'boxes': [16]}, {'id': 'U1498', 'name': 'Neon Lazers Candy Powder Filled Straws', 'total': 1, 'expiry': '2024-12-31 00:00:00', 'boxes': [1]}]}, {'begin': 486, 'end': 508, 'submitter': '1993162 Alberta Ltd. ', 'address': 'SCK4 (102 Central Ave Ste 6540)', 'name': 'Shipment 22', 'boxcount': 1, 'weightboxes': [24], 'dimensionboxes': ['20x20x20'], 'items': [{'id': 'U759', 'name': 'Nestle After Eight Advent Calendar Santa Dark Chocolate, 199g/7oz {Imported from Canada}', 'total': 12, 'expiry': '2023-03-13 00:00:00', 'boxes': [12]}, {'id': 'U841', 'name': 'Hersheys Cookies n Creme Advent Calendar 212g/ 7.4oz', 'total': 15, 'expiry': '2023-05-28 00:00:00', 'boxes': [15]}]}, {'begin': 509, 'end': 531, 'submitter': '1993162 Alberta Ltd. ', 'address': 'FTW1 (4160 Feathergrass Ln Prosper,  TX  75078)', 'name': 'Shipment 23', 'boxcount': 1, 'weightboxes': [24], 'dimensionboxes': ['14x14x14'], 'items': [{'id': 'U1344', 'name': 'Neon Tropical Candy Powder Filled Straws', 'total': 19, 'expiry': '2024-12-31 00:00:00', 'boxes': [19]}]}]
chrome_data = "C:/project/user-data"
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

url = "https://sellercentral.amazon.com/fba/sendtoamazon?ref=fbacentral_nav_fba"
driver.get(url)
try:
    check = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "div[data-testid='sku-list']")))
except:
    try:
        driver.find_element(By.CSS_SELECTOR, "input[id='signInSubmit']").click()
    except:
        input("Please click `Chrome Tester` menu, then login manually, then close the browser and try the script again")
        sys.exit()
try:
    check = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "div[data-testid='sku-list']")))
except:
    # todo: create new
    pass

defsubmitter = driver.find_element(By.CSS_SELECTOR, "div[class='textBlock-60ch break-words']").text
dlist = datalist[0]

submitter = dlist['submitter'].split("(")[0].strip()
addresstmp = dlist['address']
addresslist = addresstmp[addresstmp.find("(")+1:addresstmp.find(")")].strip().split(" ")
address = addresslist[0] + " " + addresslist[1]# + " " + addresslist[2]
print(submitter, address)
if defsubmitter.find(submitter) != -1 and defsubmitter.find(address) != -1:
    pass
else:
    driver.find_element(By.CSS_SELECTOR, "a[data-testid='ship-from-another-address-link']").click()
    ck = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "div[class='selected-address-tile']")))
    selects = driver.find_elements(By.CSS_SELECTOR, "div[class='address-tile']")
    for sel  in selects:
        txt = sel.find_element(By.CSS_SELECTOR, "div[class='tile-address']").text
        # print(txt)
        if txt.find(submitter) != -1 and txt.find(address) != -1:
            sel.find_element(By.CSS_SELECTOR, "button[class='secondary']").click()
            # input('pause')
            break
    defsubmitter = driver.find_element(By.CSS_SELECTOR, "div[class='textBlock-60ch break-words']").text
    driver.find_element(By.CSS_SELECTOR, "kat-dropdown[unique-id='katal-id-20']").click()
    driver.find_element(By.CSS_SELECTOR, "kat-dropdown[unique-id='katal-id-20']").find_element(By.CSS_SELECTOR, "div[class='select-options']").find_element(By.CSS_SELECTOR, "div[class='option-inner-container']").find_element(By.CSS_SELECTOR, "div[data-value='MSKU']").click()
    item = dlist['items'][0]
    driver.find_element(By.CSS_SELECTOR, "input[id='katal-id-21']").clear()
    xlssku = item['id'].upper()
    print('searching', xlssku, '..')
    driver.find_element(By.CSS_SELECTOR, "input[id='katal-id-21']").send_keys(xlssku)
    driver.find_element(By.CSS_SELECTOR, "a[data-testid='search-input-link']").click()
    cols = driver.find_elements(By.CSS_SELECTOR, "div[data-testid='sku-row-information-details']")
    try:
        sku = cols[0].find_element(By.CSS_SELECTOR, "div[data-testid='msku']").find_element(By.CSS_SELECTOR, "span").text
    except:
        print(xlssku, "not found!")
    if xlssku != sku:
        print(sku, "not found!")
    else:
        print(xlssku, 'has Found')
 
cols[0].find_element(By.CSS_SELECTOR, "kat-input[data-testid='sku-readiness-number-of-units-input']").find_element(By.CSS_SELECTOR, "input[name='numOfUnits']").send_keys(item['total'])
try:                
    expiry = "{}/{}/{}".format(item['expiry'][5:7], item['expiry'][8:10],item['expiry'][0:4]) 
    cols[0].find_element(By.CSS_SELECTOR, "kat-date-picker[id='expirationDatePicker']").find_element(By.CSS_SELECTOR, "input").send_keys(expiry)
    cols[0].find_element(By.CSS_SELECTOR, "kat-date-picker[id='expirationDatePicker']").find_element(By.CSS_SELECTOR, "input").send_keys(Keys.TAB)
    cols[0].find_element(By.CSS_SELECTOR, "kat-button[data-testid='skureadiness-confirm-button']").find_element(By.CSS_SELECTOR, "button[class='primary']").click()
except:
    pass
try:
    error = cols[0].find_element(By.CSS_SELECTOR, "kat-label[data-testid='sku-readiness-expiration-date-error']").text
    input(error)
except:
    pass    



options = webdriver.ChromeOptions()
options.add_argument("user-data-dir={}".format(chrome_user_data)) 
options.add_argument("profile-directory={}".format(chrome_profile))
options.add_argument('--no-sandbox')
options.add_argument("--log-level=3")
# options.add_argument("--window-size=1200, 900")
options.add_argument('--start-maximized')
options.add_argument("--disable-notifications")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)
