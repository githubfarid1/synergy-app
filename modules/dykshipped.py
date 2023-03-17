import settings
import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
from random import randint

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager as CM
from pathlib import Path
import warnings
import argparse
import os
def parse(fileoutput, chrome_data):
    warnings.filterwarnings("ignore", category=UserWarning) 
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
    print('File Selected:', fileoutput)


    driver.get('https://www2.dykpost.com/account/submitmanifest.php')
    cookies = {}
    for cookie in driver.get_cookies():
        cookies[cookie['name']] = cookie['value']


    headers = {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'Accept-Language': 'en-US,en;q=0.9',
        'Cache-Control': 'max-age=0',
        'Connection': 'keep-alive',
        # Requests sorts cookies= alphabetically
        # 'Cookie': '_ga=GA1.2.925081705.1657077797; __stripe_mid=f165f455-642d-4b54-b418-88f82def98227388a0; PHPSESSID=76remnpfr3sf72neni6qf7uuj0; __stripe_sid=a462db58-64f3-442f-8996-c4cd92486e50c79e73',
        'Referer': 'https://www2.dykpost.com/account/submitmanifest.php',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-User': '?1',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36',
        'sec-ch-ua': '".Not/A)Brand";v="99", "Google Chrome";v="103", "Chromium";v="103"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
    }

    #tes connection
    params = {
        'is': 'shipped',
        'pn': 1,
    }

    response = requests.get('https://www2.dykpost.com/account/submitmanifest.php', params=params, cookies=cookies, headers=headers)
    html = response.content
    soup = BeautifulSoup(html, "html.parser")
    trs = soup.find_all('tr', class_='ppb_text')
    if len(trs) == 0:
        input('Please, Login it first, then press any key to continue....')
        # exit()


    fres = []
    page = 0
    while True:
        page += 1
        params = {
            'is': 'shipped',
            'pn': page,
        }

        response = requests.get('https://www2.dykpost.com/account/submitmanifest.php', params=params, cookies=cookies, headers=headers)

        html = response.content

        soup = BeautifulSoup(html, "html.parser")
        islast = soup.find('ul', class_='pagination').find('a', string='Next')
        # print(islast)
        # break
        trs = soup.find_all('tr', class_='ppb_text')
        for index1, tr in enumerate(trs):
            tds = tr.find_all('td')
            no = ''
            name = ''
            status = ''
            address = ''
            status = ''
            description = ''
            origin = ''
            weight = ''
            price = ''
            tracking_number = ''
            loaded_cross = ''
            loaded_time =''

            for index2,td in enumerate(tds):
                # print(index2)
                if index2 == 0:
                    no = td.text
                if index2 == 1:
                    name = td.find('input', {'id':'recipient_name{}'.format(index1)})['value']
                
                if index2 == 2:
                    address = td.find('input', {'id':'address_line1_{}'.format(index1)})['value']
                    address += td.find('input', {'id':'address_line2_{}'.format(index1)})['value']
                    address += td.find('input', {'id':'address_line3_{}'.format(index1)})['value']
                    # address += '\n'
                    address += td.find('input', {'id':'city{}'.format(index1)})['value']
                    address += ' ' + td.find('input', {'id':'state{}'.format(index1)})['value']
                    address += ' ' + td.find('input', {'id':'zip{}'.format(index1)})['value']
                    address += ' ' + td.find('input', {'id':'country{}'.format(index1)})['value']

                if index2 == 3:
                    status = td.text.strip()
                    # print(status)
                if index2 == 4:
                    description = td.find('input', {'id':'description{}'.format(index1)})['value']
                if index2 == 5:
                    origin = td.find('input', {'id':'country_of_origin{}'.format(index1)})['value']

                if index2 == 6:
                    weight = td.find('input', {'id':'weight{}'.format(index1)})['value']

                if index2 == 7:
                    price = td.find('input', {'id':'price{}'.format(index1)})['value']

                if index2 == 8:
                    tracking_number = td.find('input', {'id':'tracking{}'.format(index1)})['value']

                if index2 == 9:
                    loaded_cross = td.find('input', {'id':'loadedshipdate{}'.format(index1)})['value']
                if index2 == 10:

                    loaded_time = td.find('input', {'id':'loadedtimestamp{}'.format(index1)})['value']


            # print(no,name, address, status, description, origin, weight, price, tracking_number, loaded_cross, loaded_time)
            dict = {
                # 'NO': no,
                'NAME': name,
                'ADDRESS': address,
                'STATUS': status,
                'ITEM DESCRIPTION': description,
                'ORIGIN': origin,
                'LB': weight,
                'USD': price,
                'TRACKING NUMBER': tracking_number,
                'LOADED TO CROSS ON': loaded_cross,
                'LOADED TIME STAMP': loaded_time
            }
                
            fres.append(dict)
        print('page', page, 'scraped')
        # time.sleep(randint(1,5))
        time.sleep(1)
        if islast == None:
            break
    df = pd.DataFrame(fres)
    df.to_excel(fileoutput , index=False)

def main():
    parser = argparse.ArgumentParser(description="DYK Shipped")
    parser.add_argument('-o', '--output', type=str,help="File Output")
    parser.add_argument('-d', '--data', type=str,help="Chrome User Data Directory")
    args = parser.parse_args()
    if args.output[-5:] != '.xlsx':
        print('File input have to XLSX file')
        exit()
    isExist = os.path.exists(args.data)
    if isExist == False :
        print('Please check Chrome User Data Directory')
        exit()
    isExist = os.path.exists(args.output)
    if isExist == False :
        print('Please check XLSX file')
        exit()
    
    parse(args.output, args.data)
    
if __name__ == '__main__':
    main()
