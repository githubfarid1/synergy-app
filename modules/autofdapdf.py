from single_fdaentry import FdaEntry
from single_fdapdf import FdaPdf
import argparse
import sys
from sys import platform
import os
import shutil
import time
import fitz
import unicodedata as ud
import uuid
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager as CM
import warnings
from random import randint
import glob
import string
from datetime import date
import json
import amazon_lib as lib
import xlwings as xw
import logging


warnings.filterwarnings("ignore", category=Warning)
logger = logging.getLogger()
logger.setLevel(logging.NOTSET)

def explicit_wait():
    time.sleep(randint(1, 3))

def getConfig():
	file = open("setting.json", "r")
	config = json.load(file)
	return config

def format_filename(s):
    valid_chars = "-_.() %s%s" % (string.ascii_letters, string.digits)
    filename = ''.join(c for c in s if c in valid_chars)
    filename = filename.replace(' ','_') # I don't like spaces in filenames.
    return filename

def deltree(folder):
    print("Trying removing", folder, "Folder ...", end=" ", flush=True)
    try:
        shutil.rmtree(folder)
    except OSError as e:
        result = "Error: %s : %s" % (folder, e.strerror)
    result =  "Success"
    print(result)

def pdf_rename(pdfoutput_folder):
    pdffolder = pdfoutput_folder
    delimeter = file_delimeter()
    # print("Renaming Files started")
    list_of_files = glob.glob(pdffolder + delimeter + "*.pdf" )
    exceptedfiles = []
    for file in list_of_files:
        if file.find("filename") != -1:
            exceptedfiles.append(file)
    if len(exceptedfiles) == 0:
        return ""
     
    latest_file = max(exceptedfiles, key=os.path.getctime)
    filename = latest_file
    if platform == "win32":
        filename = filename.split("\\")[-1]
    else:
        filename = filename.split("/")[-1]
           
    doc = fitz.open(pdffolder + delimeter + filename)
    page = doc[0]
    search = page.get_text("blocks", clip=[100.6500015258789, 271.04034423828125, 185.60845947265625, 283.09893798828125])
    tmpname = search[0][4].replace(".", "")
    strdate = str(date.today())
    pdfsubmitter = format_filename("{}_{}.{}".format(tmpname, strdate, "pdf"))
    doc.close()
    isExist = os.path.exists(pdffolder + delimeter + pdfsubmitter)
    if isExist:
        os.remove(pdffolder + delimeter + pdfsubmitter)

    print("rename", pdffolder + delimeter + filename)
    os.rename(pdffolder + delimeter + filename, pdffolder + delimeter + pdfsubmitter)
    return pdfsubmitter

def webentry_update(pdffile,  pdffolder):
    print("Update Web Entry Identification Started..")
    time.sleep(1)
    delimeter = file_delimeter()
    doc = fitz.open(pdffolder + delimeter + pdffile)
    page = doc[0]
    submitter = page.get_text("block", clip=[100.6500015258789, 271.04034423828125, 185.60845947265625, 283.09893798828125]).strip()
    entry_id = page.get_text("block", clip=(152.7100067138672, 202.04034423828125, 230.7493438720703, 214.09893798828125)).strip()

    # print(submitter, entry_id)
    for i in range(2, MAXROW+2):
        if xlworksheet['B{}'.format(i)].value == None:
            break
        if xlworksheet['T{}'.format(i)].value.strip() == submitter:
            xlworksheet['A{}'.format(i)].value = entry_id
    # workbook.save(xlsfilename)
    print(submitter, "Updated")
    time.sleep(1)

def browser_init(chrome_data, pdfoutput_folder):
    warnings.filterwarnings("ignore", category=UserWarning)
    config = getConfig()
    options = webdriver.ChromeOptions()
    # options.add_argument("--headless")
    # options.add_argument("user-data-dir={}".format(chrome_data)) #Path to your chrome profile
    options.add_argument("user-data-dir={}".format(config['chrome_user_data'])) 
    options.add_argument("profile-directory={}".format(config['chrome_profile']))
    options.add_argument('--no-sandbox')
    options.add_argument("--log-level=3")
    # options.add_argument("--window-size=1200, 900")
    options.add_argument('--start-maximized')
    options.add_argument("--disable-notifications")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    download_dir = pdfoutput_folder
    profile = {"plugins.plugins_list": [{"enabled": False, "name": "Chrome PDF Viewer"}], # Disable Chrome's PDF Viewer
                "download.default_directory": download_dir, 
                "download.extensions_to_open": "applications/pdf", 
                'profile.default_content_setting_values.automatic_downloads': 1,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "plugins.always_open_pdf_externally": True #It will not show PDF directly in chrome                    
            }
    options.add_experimental_option("prefs", profile)
    return webdriver.Chrome(service=Service(CM().install()), options=options)

def browser_login(driver):
    loginurl = "https://www.access.fda.gov/oaa/logonFlow.htm?execution=e1s1"
    # driver = self.__driver
    driver.get(loginurl)
    WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input[id='understand']")))
    driver.find_element(By.CSS_SELECTOR, "input[id='understand']").click()
    explicit_wait()
    driver.find_element(By.CSS_SELECTOR, "a[id='login']").click()
    explicit_wait()
    driver.find_element(By.CSS_SELECTOR, "a[title='Prior Notice System Interface']").click()
    explicit_wait()
    driver.find_element(By.CSS_SELECTOR, "img[alt='Create New Web Entry']").click()
    explicit_wait()
    return driver

def clear_screan():
    return
    if platform == "win32":
        os.system("cls")
    else:
        os.system("clear")

def file_delimeter():
    delimeter = "/"    
    if platform == "win32":
        delimeter = "\\"
    return delimeter

def clearlist(*args):
    for varlist in args:
        varlist.clear()

def del_non_annot_page(pdffiles, pdffolder):
    print("Removing Non Highlight Pages..")
    tmpfile = pdffolder + file_delimeter() + "tmp.pdf"
    for pdffile in pdffiles:
        shutil.copy(pdffile, tmpfile)
        doc = fitz.open(pdffolder + file_delimeter() + "tmp.pdf")
        selected = []
        for idx, page in enumerate(doc):
            for annot in page.annots():
                selected.append(idx)
                break
        selected.append(0)
        selected = list(dict.fromkeys(selected))
        selected.sort()
        doc.select(selected)
        doc.save(pdffile)
        print(os.path.basename(pdffile), "passed.")
        doc.close()
        time.sleep(1)
    isExist = os.path.exists(tmpfile)
    
    if isExist:    
        os.remove(tmpfile)    
    print("")

def join_folderpdf(pdffiles, pdfoutput_folder):
    print("Merging PDF files in one folder started..")
    time.sleep(1)

    foldername = pdfoutput_folder + file_delimeter() + "combined"
    isExist = os.path.exists(foldername)
    if isExist:
        deltree(foldername)
    else:
        os.makedirs(foldername)

    dictfiles = {}
    for pdffile in pdffiles:
        basefilename = os.path.basename(pdffile)
        dictfiles[int(basefilename.replace(".pdf",""))] = pdffile
    sortedfiles = dict(sorted(dictfiles.items()))

    for file in sortedfiles:
        print(os.path.basename(sortedfiles[file]), "merged")
        time.sleep(1)
        shutil.move(sortedfiles[file], foldername + file_delimeter())
    print("Merging PDF files finished..")

def xls_dataframe_generator(filename, sname):
    df = pd.read_excel(filename, sheet_name=sname)
    cols = df.groupby('Shiper Address').first().values.tolist()
    print(cols)

def xls_data_generator(xlws, maxrow):
    global xlworksheet
    global MAXROW
    xlworksheet = xlws #xlworkbook.sheets[sname]
    MAXROW = maxrow
    allData = {}
    wcode = []
    wshipper = []
    wdesc = []
    wsize = []
    wtotal = []
    wmanufact = []
    wmanufact_addr = []
    wmanufact_city = []
    wconsignee = []
    wconsignee_addr = []
    wconsignee_city = []
    wconsignee_postal = []
    wconsignee_state = []
    wconsignee_stact = []
    wsubmitter = []
    wsubmitter_add = []
    wsubmitter_cityetc = []
    wsubmitter_country = []
    wpnumber = []
    wbox = []
    wentrycode = []
    wsku = []
    wentryid = xlworksheet['B{}'.format(2)].value
    for i in range(2, MAXROW+2):
        if wentryid != xlworksheet['B{}'.format(i)].value:# and xlworksheet['B{}'.format(i)].value != None:
            rid = uuid.uuid4().hex
            allData[rid] = {'data':list(zip(wshipper, wcode, wdesc, wsize, wtotal, wmanufact, wmanufact_addr, wmanufact_city, wconsignee, wconsignee_addr, wconsignee_city, wconsignee_postal, wconsignee_stact, wconsignee_state, wsubmitter, wsubmitter_add, wsubmitter_cityetc, wsubmitter_country, wpnumber, wbox, wentrycode, wsku)),
            'count' : len(wcode)} 
            wentryid = xlworksheet['B{}'.format(i)].value
            clearlist(wshipper, wcode, wdesc, wsize, wtotal, wmanufact, wmanufact_addr, wmanufact_city, wconsignee, wconsignee_addr, wconsignee_city, wconsignee_postal, wconsignee_stact, wconsignee_state, wsubmitter, wsubmitter_add, wsubmitter_cityetc, wsubmitter_country, wpnumber, wbox, wentrycode, wsku)
        
        if xlworksheet['B{}'.format(i)].value == None:
            break
        
        wshipper.append(str(xlworksheet['B{}'.format(i)].value).strip())
        wcode.append(str(xlworksheet['F{}'.format(i)].value).strip())
        strdesc= ud.normalize('NFKD', str(xlworksheet['G{}'.format(i)].value).strip()).encode('ascii', 'ignore').decode('ascii')
        wdesc.append(strdesc)
        wsize.append(str(xlworksheet['H{}'.format(i)].value).strip())
        wtotal.append(str(xlworksheet['I{}'.format(i)].value).strip())
        strmanufact = ud.normalize('NFKD', str(xlworksheet['K{}'.format(i)].value).strip()).encode('ascii', 'ignore').decode('ascii')
        wmanufact.append(strmanufact)
        strmanufact_addr = ud.normalize('NFKD', str(xlworksheet['L{}'.format(i)].value).strip()).encode('ascii', 'ignore').decode('ascii')
        wmanufact_addr.append(strmanufact_addr)
        wmanufact_city.append(str(xlworksheet['M{}'.format(i)].value).strip())
        wconsignee.append(str(xlworksheet['N{}'.format(i)].value).strip())
        wconsignee_addr.append(str(xlworksheet['O{}'.format(i)].value).strip())
        wconsignee_city.append(str(xlworksheet['P{}'.format(i)].value).strip())
        wconsignee_postal.append(str(xlworksheet['Q{}'.format(i)].value).strip())
        wconsignee_state.append(str(xlworksheet['R{}'.format(i)].value).strip())
        wconsignee_stact.append(str(xlworksheet['S{}'.format(i)].value).strip())
        wsubmitter.append(str(xlworksheet['T{}'.format(i)].value).strip())
        wsubmitter_add.append(str(xlworksheet['U{}'.format(i)].value).strip())
        wsubmitter_cityetc.append(str(xlworksheet['V{}'.format(i)].value).strip())
        wsubmitter_country.append(str(xlworksheet['W{}'.format(i)].value).strip())
        wpnumber.append("")
        wbox.append(str(xlworksheet['D{}'.format(i)].value).strip())
        wentrycode.append(str(xlworksheet['A{}'.format(i)].value).strip())
        wsku.append(str(xlworksheet['E{}'.format(i)].value).strip())
    return allData

def choose_pdf_file(file_list, entry_id):
    for file in file_list:
        doc = fitz.open(file)
        page = doc[0]
        ex_entry_id = page.get_text("block", clip=(152.7100067138672, 202.04034423828125, 230.7493438720703, 214.09893798828125)).strip()
        if entry_id == ex_entry_id:
            return file
    return ""
    
def save_to_xls(pnlist):
    for i in range(2, MAXROW+2):
        # strdesc = ud.normalize('NFKD', str(worksheet['G{}'.format(i)].value).strip()).encode('ascii', 'ignore').decode('ascii')
        sku = xlworksheet['E{}'.format(i)].value
        if sku == None:
            break
        for pn in pnlist:
            if xlworksheet['A{}'.format(i)].value == pn['entry_id'] and sku == pn['sku'] and xlworksheet['D{}'.format(i)].value == pn['boxes']:
                    xlworksheet['X{}'.format(i)].value = "'" + pn['pnnumber']
                    break
    # try:        
    #     workbook.save(filename)
    # except:
    #     input("Save to excel Failed!!. Make sure you have closed it. Run the script again.")
    #     sys.exit()


def main():
    parser = argparse.ArgumentParser(description="FDA Entry + PDF Extractor")
    parser.add_argument('-i', '--input', type=str,help="File Input")
    parser.add_argument('-s', '--sheet', type=str,help="Sheet Name")
    parser.add_argument('-dt', '--date', type=str,help="Arrival Date")
    parser.add_argument('-d', '--chromedata', type=str,help="Chrome User Data Directory")
    parser.add_argument('-o', '--output', type=str,help="PDF output folder")
    
    args = parser.parse_args()
    if not (args.input[-5:] == '.xlsx' or args.input[-5:] == '.xlsm'):
        input('input the right XLSX or XLSM file')
        sys.exit()
    isExist = os.path.exists(args.chromedata)
    if isExist == False :
        input('Please check Chrome User Data Directory')
        sys.exit()
    isExist = os.path.exists(args.input)
    if isExist == False :
        input('Please check XLSX or XLSM file')
        sys.exit()
    if len(args.date) != 10:
        input('Date Arrival is wrong')
        sys.exit()

    isExist = os.path.exists(args.output)
    if isExist == False :
        input('Please make sure PDF folder is exist')
        sys.exit()
    isExist = os.path.exists(args.chromedata)
    if isExist == False :
        input('Please check Chrome User Data Directory')
        sys.exit()

    clear_screan()
    file_handler = logging.FileHandler('logs/autofda-err.log')
    file_handler.setLevel(logging.ERROR)
    file_handler_format = '%(asctime)s | %(levelname)s | %(lineno)d: %(message)s'
    file_handler.setFormatter(logging.Formatter(file_handler_format))
    logger.addHandler(file_handler)

    print('Creating Excel Backup File...', end="", flush=True)
    fnameinput = os.path.basename(args.input)
    pathinput = args.input[0:-len(fnameinput)]
    backfile = "{}{}_backup{}".format(pathinput, os.path.splitext(fnameinput)[0], os.path.splitext(fnameinput)[1])
    shutil.copy(args.input, backfile)
    print('OK')
    print(args.input)

    print('Opening the Source Excel File...', end="", flush=True)
    xlbook = xw.Book(args.input)
    print('OK')
    # input(json.dumps(xlsdictwcode))
    # exit()
    strdate = str(date.today())
    foldernamepn = "{}{}_{}".format(args.output + lib.file_delimeter(), 'prior_notice', strdate) 
    isExist = os.path.exists(foldernamepn)
    if not isExist:
        os.makedirs(foldernamepn)
    
    xlsfilename = os.path.basename(args.input)
    foldername = format_filename("{}_{}_{}".format(xlsfilename.replace(".xlsx", "").replace(".xlsm", ""), args.sheet, strdate) )
    
    complete_output_folder = foldernamepn + file_delimeter() + foldername
    isExist = os.path.exists(complete_output_folder)
    if not isExist:
        os.makedirs(complete_output_folder)
    maxrun = 10
    xlsheet = xlbook.sheets[args.sheet]
    for i in range(1, maxrun+1):
        if i > 1:
            print("Error, Process will be reapeated - {} trying".format(i-1) )
        try:
            maxrow = xlsheet.range('B' + str(xlsheet.cells.last_cell.row)).end('up').row
            xlsdictall = xls_data_generator(xlws=xlsheet, maxrow=maxrow)
            xlsdictwcode = {}
            for idx, xls in xlsdictall.items():
                for data in xls['data']:
                    if data[20] == 'None':
                        xlsdictwcode[idx] = xls
                        break

            for xlsdata in xlsdictwcode.values():
                
                try:
                    driver.close()
                    driver.quit()
                except:
                    pass
                driver = browser_init(chrome_data=args.chromedata, pdfoutput_folder=complete_output_folder)
                driver = browser_login(driver)

                fda_entry = FdaEntry(driver=driver, datalist=xlsdata, datearrival=args.date, pdfoutput=complete_output_folder)
                # if not first:
                try:
                    driver.find_element(By.CSS_SELECTOR, "img[alt='Create WebEntry Button']").click()
                except:
                    pass
                fda_entry.parse()
                # time.sleep(2)
                pdf_filename = pdf_rename(pdfoutput_folder=complete_output_folder)
                if pdf_filename != "":
                    webentry_update(pdffile=pdf_filename, pdffolder=complete_output_folder)
                    try:
                        xlbook.save(args.input)
                    except:
                        pass
                    

                else:
                    print("file:", pdf_filename)
                    input("rename the file was failed")
                # print("sdsds")
                first = False
            
            #regenerate data
            xlsdictall = xls_data_generator(xlws=xlsheet, maxrow=maxrow)
            
            #Keep only submitter PDF files.
            submitters = []
            for xlsdata in xlsdictall.values():
                submitters.append(format_filename(xlsdata['data'][0][14].replace(".", "")))
            all_pdf_files = glob.glob(complete_output_folder + file_delimeter() + "*.pdf")
            for file in all_pdf_files:
                found = False
                for submitter in submitters:
                    if file.find(submitter) != -1:
                        found = True
                        break
                if not found:
                    os.remove(file)
            dirs = [d for d in os.listdir(complete_output_folder) if os.path.isdir(os.path.join(complete_output_folder, d))]
            for dir in dirs:
                deltree(complete_output_folder + file_delimeter() + dir)

            list_of_files = glob.glob(complete_output_folder + file_delimeter() + "*.pdf")
            allsavedfiles = []
            for xlsdata in xlsdictall.values():
                entry_id = xlsdata['data'][0][20]
                pdf_filename = choose_pdf_file(list_of_files, entry_id)
                print('PDF File processing: ', pdf_filename)
                prior = FdaPdf(filename=pdf_filename, datalist=xlsdata, pdfoutput=complete_output_folder)
                prior.highlightpdf_generator()
                prior.insert_text()
                save_to_xls(pnlist=prior.pnlist)
                try:
                    xlbook.save(args.input)
                except:
                    pass

                allsavedfiles.extend(prior.savedfiles)
            
            setall = set(allsavedfiles)

            if len(setall) != len(allsavedfiles):
                input("Combining all pdf files Failed because there are one or more files is has the same name.")
            else:
                del_non_annot_page(allsavedfiles, complete_output_folder)
                join_folderpdf(allsavedfiles, complete_output_folder)
                resultfile = lib.join_pdfs(source_folder=complete_output_folder + lib.file_delimeter() + "combined", output_folder=complete_output_folder, tag="FDA_All")
                print(resultfile, "created")
            break
        except Exception as e:
            logger.error(e)
            print("There is an error, check logs/autofda-err.log")
            try:
                xlbook.save(args.input)
            except:
                pass
            # xlbook.close()
            if i == maxrun:
                logger.error("Execution Limit reached, Please check the script")
            continue
    # Delete all file folder
    for filename in list_of_files:
        folder = filename[:-4]
        deltree(folder)
    input("data generating completed...")


if __name__ == '__main__':
    main()
