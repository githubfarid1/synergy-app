# import settings
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.select import Select
from selenium.webdriver.support import expected_conditions as EC
import time
import os
from random import randint
# from random import rand

import string
from sys import platform
import numpy as np

'''PROBLEM:
 RUNNING THE SCRIPT ON WINDOWS TERMINAL MAKE PAUSE SUDDENLY
 SOLUTION: DISABLE QUICK EDIT MODE IN TERMINAL
 https://stackoverflow.com/questions/73486528/python-script-pausing-in-cmd
'''
def explicit_wait():
    # time.sleep(randint(1, 2))
    time.sleep(np.random.randint(1, 2))
    
    # return

def explicit_wait_ext():
    # time.sleep(randint(3, 5))
    time.sleep(np.random.randint(2, 4))
    # return

def clearlist(*args):
    for varlist in args:
        varlist.clear()

def clear_screan():
    if platform == "win32":
        os.system("cls")
    else:
        os.system("clear")

def pause(mess=""):
    input(mess)

def format_filename(s):
    valid_chars = "-_.() %s%s" % (string.ascii_letters, string.digits)
    filename = ''.join(c for c in s if c in valid_chars)
    filename = filename.replace(' ','_') # I don't like spaces in filenames.
    return filename

class FdaEntry:
    def __init__(self, driver, datalist, datearrival, pdfoutput) -> None:
        print("Initialated FDA Entry..")
        time.sleep(1)
        self.__datearrival = datearrival
        self.__datalist = datalist
        self.__pdfoutput = pdfoutput
        self.__savedfiles = []
        self.__pdffilelist = []
        self.__driver = driver
        self.__delimeter = "/"    
        if platform == "win32":
            self.__delimeter = "\\"
        self.__download_filename = ''        
    @property
    def datearrival(self):
        return self.__datearrival

    @datearrival.setter
    def datearrival(self, value):
        ldate = value.split("-")
        strdate = "{}/{}/{}".format(ldate[1], ldate[2], ldate[0])
        self.__datearrival.set(strdate)

    @property
    def xlsname(self):
        return self.__xlsname

    @xlsname.setter
    def xlsname(self, value):
        self.__xlsname = value

    @property
    def download_filename(self):
        return self.__download_filename

    @download_filename.setter
    def download_filename(self, value):
        self.__download_filename = value

    @property
    def driver(self):
        return self.__driver

    @property
    def datalist(self):
        return self.__datalist

    @datalist.setter
    def datalist(self, value):
        self.__datalist = value

    @property
    def pdffilelist(self):
        return self.__pdffilelist

    @pdffilelist.setter
    def pdffilelist(self, value):
        self.__pdffilelist = value

    @property
    def pdfpage(self):
        return self.__pdfpage

    @property
    def savedfiles(self):
        return self.__savedfiles

    @property
    def pdfoutput_folder(self):
        return self.__pdfoutput

    @property
    def file_delimeter(self):
        return self.__delimeter

    def parse(self):
        datatable = self.datalist
        datearrival = self.datearrival
        driver = self.driver
        # SET WEB ENTRY
        pncount = str(datatable['count'])
        print("PN Web Entry", datatable['data'][0][14], "Started.. ", "(" + pncount + " Products)")
        Select(driver.find_element(By.CSS_SELECTOR, "select[name='webEntry.entryType.code']")).select_by_visible_text('Consumption')
        explicit_wait()
        driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")
        button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "img[alt='Next Button']")))
        button.click()
        explicit_wait()
        button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[name='generateIdFlag']")))
        button.click()
        explicit_wait()
        driver.find_element(By.CSS_SELECTOR, "input[name='webEntry.intendedPNCount']").send_keys(pncount)
        explicit_wait()
        driver.find_element(By.CSS_SELECTOR, "input[name='webEntry.portOfArrival.portCode']").send_keys("3310")
        explicit_wait()
        ldate = datearrival.split("-")
        strdate = "{}/{}/{}".format(ldate[1], ldate[2], ldate[0])
        driver.find_element(By.CSS_SELECTOR, "input[name='anticipatedArrivalDate']").send_keys(strdate)
        explicit_wait()
        Select(driver.find_element(By.CSS_SELECTOR, "select[name='hourValue']")).select_by_visible_text('09')
        time.sleep(1)
        Select(driver.find_element(By.CSS_SELECTOR, "select[name='minValue']")).select_by_visible_text('00')
        explicit_wait()
        Select(driver.find_element(By.CSS_SELECTOR, "select[name='submitterSameAsRoleId']")).select_by_visible_text('Yes')
        explicit_wait()
        Select(driver.find_element(By.CSS_SELECTOR, "select[name='importerSameAsRoleId']")).select_by_visible_text('Yes')
        explicit_wait()
        button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "img[alt='Enter Submitter Button']")))
        button.click()
        explicit_wait()
        # added
        # input('')
        button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[name='submitterSameAsRoleId']")))
        button.click()
        explicit_wait()
        wsubmitter = datatable['data'][0][14]
        wsubmitter_add = datatable['data'][0][15]
        wsubmitter_cityetc = datatable['data'][0][16]
        wsubmitter_clist = wsubmitter_cityetc.split("/")
        
        driver.find_element(By.CSS_SELECTOR, "input[name='submitter.name']").clear() 
        driver.find_element(By.CSS_SELECTOR, "input[name='submitter.name']").send_keys(wsubmitter)
        explicit_wait()
        driver.find_element(By.CSS_SELECTOR, "input[name='submitter.address.address1']").clear()
        driver.find_element(By.CSS_SELECTOR, "input[name='submitter.address.address1']").send_keys(wsubmitter_add)
        explicit_wait()
        driver.find_element(By.CSS_SELECTOR, "input[name='submitter.address.city']").clear()
        driver.find_element(By.CSS_SELECTOR, "input[name='submitter.address.city']").send_keys(wsubmitter_clist[0])
        explicit_wait()
        Select(driver.find_element(By.CSS_SELECTOR, "select[id='requiring work']")).select_by_value("{}-{}".format("CA", wsubmitter_clist[1]))
        explicit_wait()
        driver.find_element(By.CSS_SELECTOR, "input[name='submitter.address.zipMailCode']").clear()
        driver.find_element(By.CSS_SELECTOR, "input[name='submitter.address.zipMailCode']").send_keys(wsubmitter_clist[2])
        explicit_wait()
        #---

        button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "img[alt='Save Button']")))
        button.click()
        explicit_wait()
        button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[name='useNewAddr'][value='1']")))
        button.click()
        explicit_wait()
        button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "img[alt='OK Button']")))
        button.click()
        explicit_wait()
        Select(driver.find_element(By.CSS_SELECTOR, "select[name='motCode']")).select_by_visible_text('Land, Truck')
        explicit_wait()
        button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "img[alt='Enter Carrier Button']")))
        button.click()
        explicit_wait()
        driver.find_element(By.CSS_SELECTOR, "input[name='carrier.name']").send_keys("DYKP")
        explicit_wait()
        button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "img[alt='Save Button']")))
        button.click()
        explicit_wait()
        button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "img[alt='Save Button']")))
        button.click()
        explicit_wait()

        for data in datatable['data']:
            # try:
            button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "img[alt='Create PN Button']")))
            button.click()
            explicit_wait()
            Select(driver.find_element(By.CSS_SELECTOR, "select[name='shippingCountryCode']")).select_by_visible_text('Canada  (CA)')
            explicit_wait()
            wcode = data[1]
            driver.find_element(By.CSS_SELECTOR, "input[name='pnArticle.product.fdaProductCode']").send_keys(wcode)
            explicit_wait()
            wdesc = data[2]
            driver.find_element(By.CSS_SELECTOR, "input[name='pnArticle.product.productCommonName']").send_keys(wdesc)
            explicit_wait()
            wsize = data[3]
            driver.find_element(By.CSS_SELECTOR, "input[name='pnArticle.baseUnitNumber']").send_keys(wsize)
            explicit_wait()
            Select(driver.find_element(By.CSS_SELECTOR, "select[name='uomCode']")).select_by_visible_text('Grams')
            explicit_wait()
            wtot = data[4]
            driver.find_element(By.CSS_SELECTOR, "input[name='pnArticle.packageItem0.containerQty']").send_keys(wtot)
            explicit_wait()
            Select(driver.find_element(By.CSS_SELECTOR, "select[name='containerTypeCode0']")).select_by_visible_text('Box')
            explicit_wait()
            button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "img[alt='Save Button']")))
            button.click()
            explicit_wait()
            select = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "select[name='pnArticle.pnFacilities.producer.address.country.countryCode']")))
            Select(select).select_by_visible_text('Canada  (CA)')
            explicit_wait()
            button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "img[alt='Enter Manufacturer Button']")))
            button.click()            
            explicit_wait()
            wmanuc = data[5]
            driver.find_element(By.CSS_SELECTOR, "input[name='pnArticle.pnFacilities.producer.name']").send_keys(wmanuc)
            explicit_wait()
            wmanuc_addr = data[6]
            driver.find_element(By.CSS_SELECTOR, "input[name='pnArticle.pnFacilities.producer.address.address1']").send_keys(wmanuc_addr)
            explicit_wait()
            wmanuc_city = data[7]
            driver.find_element(By.CSS_SELECTOR, "input[name='pnArticle.pnFacilities.producer.address.city']").send_keys(wmanuc_city)
            explicit_wait()
            button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[name='pnArticle.pnFacilities.producer.regExemptFlag']")))
            button.click()            
            explicit_wait()
            Select(driver.find_element(By.CSS_SELECTOR, "select[id='requiring work']")).select_by_value("11")
            explicit_wait()
            button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "img[alt='Save Button']")))
            button.click()
            explicit_wait()
            Select(driver.find_element(By.CSS_SELECTOR, "select[name='pnArticle.pnFacilities.shipper.address.country.countryCode']")).select_by_visible_text('Canada  (CA)')

            explicit_wait()
            button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "img[alt='Enter Shipper Button']")))
            button.click()            
            explicit_wait()

            Select(driver.find_element(By.CSS_SELECTOR, "select[id='State']")).select_by_value("8")
            explicit_wait()
            button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "img[alt='Save Button']")))
            button.click()            

            explicit_wait()
            Select(driver.find_element(By.CSS_SELECTOR, "select[name='pnArticle.pnFacilities.owner.address.country.countryCode']")).select_by_visible_text('Canada  (CA)')
            explicit_wait()
            button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "img[alt='Enter Owner Button']")))
            button.click()            
            explicit_wait()
            Select(driver.find_element(By.CSS_SELECTOR, "select[id='State']")).select_by_value("8")
            explicit_wait()
            button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "img[alt='Save Button']")))
            button.click()            
            explicit_wait()
            button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "img[alt='Enter Consignee Button']")))
            button.click()            
            explicit_wait()
            wcon = data[8]
            driver.find_element(By.CSS_SELECTOR, "input[name='pnArticle.pnFacilities.consignee.name']").send_keys(wcon)
            explicit_wait()
            wcon_addr = data[9]
            driver.find_element(By.CSS_SELECTOR, "input[name='pnArticle.pnFacilities.consignee.address.address1']").send_keys(wcon_addr)
            explicit_wait()
            wcon_city = data[10]
            driver.find_element(By.CSS_SELECTOR, "input[name='pnArticle.pnFacilities.consignee.address.city']").send_keys(wcon_city)
            explicit_wait()
            wcon_postal = data[11]
            driver.find_element(By.CSS_SELECTOR, "input[name='pnArticle.pnFacilities.consignee.address.zipMailCode']").send_keys(wcon_postal)
            explicit_wait()
            wcon_state = data[12]
            Select(driver.find_element(By.CSS_SELECTOR, "select[name='pnArticle.pnFacilities.consignee.address.subdivision.code']")).select_by_visible_text(wcon_state)
            explicit_wait()
            button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "img[alt='Save Button']")))
            button.click()            
            explicit_wait_ext()
            driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")
            button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "img[alt='PN Save Button']")))
            button.click()
            explicit_wait_ext()
            print(wcode, "Saved")
            driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")
            explicit_wait_ext()
            button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "img[alt='Next Button']")))
            button.click()
            explicit_wait()
            button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "img[alt='Cancel Button']")))
            button.click()
            explicit_wait()
            # except:
            #     input("Error Found..")
        try:
            button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "img[alt='Complete Web Entry Button']")))
            button.click()            
            explicit_wait()
            driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")
            button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "img[alt='Next Button']")))
            button.click()            
            time.sleep(np.random.randint(2, 4))
            button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "img[alt='Print Summary Button']")))
            button.click()            
            time.sleep(5)
            print("PN Web Entry", datatable['data'][0][14], "End.")
        except:
            input("Error found")


if __name__ == '__main__':
    print('This module can not run via Main')
