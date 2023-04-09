import sys
import os
import argparse
from sys import platform
from datetime import date, datetime, timedelta
import amazon_lib as lib
import amazonship
import logging
from pathlib import Path
import autofdapdf as fdaauto
from single_fdaentry import FdaEntry
from single_fdapdf import FdaPdf
from selenium.webdriver.common.by import By
import glob
import shutil
import xlwings as xw

if platform == "linux" or platform == "linux2":
    PYLOC = "python"
elif platform == "win32":
    PYLOC = "python.exe"
logger = logging.getLogger()
logger.setLevel(logging.NOTSET)
logger2 = logging.getLogger()
logger2.setLevel(logging.NOTSET)

def main():
    parser = argparse.ArgumentParser(description="Amazon Shipment")
    parser.add_argument('-xls', '--xlsinput', type=str,help="XLSX File Input")
    parser.add_argument('-shipsheet', '--shipsheet', type=str,help="Shipment Sheet of XLSX file")
    parser.add_argument('-pnsheet', '--pnsheet', type=str,help="PN Sheet of XLSX file")
    parser.add_argument('-tracksheet', '--tracksheet', type=str,help="Tracking Sheet of XLSX file")
    parser.add_argument('-output', '--pdfoutput', type=str,help="PDF output folder")
    parser.add_argument('-cdata', '--chromedata', type=str,help="Chrome User Data Directory")
    parser.add_argument('-dt', '--date', type=str,help="Arrival Date")

    args = parser.parse_args()
    if not (args.xlsinput[-5:] == '.xlsx' or args.xlsinput[-5:] == '.xlsm'):
        input('File input have to XLSX or XLSM file')
        sys.exit()
    isExist = os.path.exists(args.xlsinput)
    if not isExist:
        input(args.xlsinput + " does not exist")
        sys.exit()
    isExist = os.path.exists(args.chromedata)
    if isExist == False :
        input('Please check Chrome User Data Directory')
        sys.exit()
    isExist = os.path.exists(args.pdfoutput)
    if not isExist:
        input(args.pdfoutput + " folder does not exist")
        sys.exit()
    strdate = str(date.today())
    folderamazonship = "{}{}_{}".format(args.pdfoutput + lib.file_delimeter(), 'shipment_creation', strdate) 
    isExist = os.path.exists(folderamazonship)
    if not isExist:
        os.makedirs(folderamazonship)

    foldernamepn = "{}{}_{}".format(args.pdfoutput + lib.file_delimeter(), 'prior_notice', strdate) 
    isExist = os.path.exists(foldernamepn)
    if not isExist:
        os.makedirs(foldernamepn)



    # print("1. Shipment Creation")
    file_handler = logging.FileHandler('logs/amazonship-err.log')
    file_handler.setLevel(logging.ERROR)
    file_handler_format = '%(asctime)s | %(levelname)s | %(lineno)d: %(message)s'
    file_handler.setFormatter(logging.Formatter(file_handler_format))
    logger.addHandler(file_handler)

    file_handler2 = logging.FileHandler('logs/amazonship-info.log')
    file_handler2.setLevel(logging.INFO)
    # file_handler2_format = '%(asctime)s | %(levelname)s: %(message)s'
    file_handler2_format = '%(asctime)s | %(levelname)s | %(lineno)d: %(message)s'
    file_handler2.setFormatter(logging.Formatter(file_handler2_format))
    logger2.addHandler(file_handler2)
    fnameinput = os.path.basename(args.xlsinput)
    pathinput = args.xlsinput[0:-len(fnameinput)]
    destfile = "{}{}_new{}".format(pathinput, os.path.splitext(fnameinput)[0], os.path.splitext(fnameinput)[1])
    shutil.copy(args.xlsinput, destfile)
    xlbook = xw.Book(destfile)
    # exit()
    logger2.info("###### Start ######")
    logger2.info("Filename: {}\nSheet Name:{}\nPDF Output Folder:{}".format(destfile, args.shipsheet, folderamazonship))
    maxrun = 10
    for i in range(1, maxrun+1):
        if i > 1:
            print("Process will be reapeated")
        try:    
            shipment = amazonship.AmazonShipment(xlsfile=destfile, sname=args.shipsheet, chrome_data=args.chromedata, download_folder=folderamazonship, xlworkbook=xlbook)
            shipment.data_sanitizer()
            if len(shipment.datalist) == 0:
                break
            shipment.parse()
            shipment.xlworkbook.save(shipment.xlsfile)
            shipment.workbook.close()
        except Exception as e:
            logger.error(e)
            print("There is an error, check logs/amazonship-err.log")
            # shipment.workbook.save(shipment.xlsfile)
            # shipment.workbook.close()
            shipment.xlworkbook.save(shipment.xlsfile)
            shipment.workbook.close()
            if i == maxrun:
                logger.error("Execution Limit reached, Please check the script")
            continue
        break

    addressfile = Path("address.csv")
    resultfile = lib.join_pdfs(source_folder=folderamazonship + lib.file_delimeter() + "combined" , output_folder = folderamazonship, tag='Labels')
    print(resultfile, "created")
    if resultfile != "":
        lib.add_page_numbers(resultfile)
        lib.generate_xls_from_pdf(resultfile, addressfile)
        
    lib.copysheet(destination=destfile, source=resultfile[:-4] + ".xlsx", cols=('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'), sheetsource="Sheet", sheetdestination="Shipment labels summary", tracksheet=args.tracksheet, xlbook=xlbook)
    xlbook.save(destfile)
    # xlbook.close()
    # input("End Process..")    
    # -----------------
    xlsheet = xlbook.sheets[args.pnsheet]
    maxrow = xlsheet.range('B' + str(xlsheet.cells.last_cell.row)).end('up').row
    xlsdictall = fdaauto.xls_data_generator(xlws=xlsheet, maxrow=maxrow)
    xlsdictwcode = {}
    for idx, xls in xlsdictall.items():
        for data in xls['data']:
            if data[20] == 'None':
                xlsdictwcode[idx] = xls
                break

    xlsfilename = os.path.basename(destfile)
    strdate = str(date.today())
    foldername = fdaauto.format_filename("{}_{}_{}".format(xlsfilename[:-5], args.pnsheet, strdate) )
    complete_output_folder = foldernamepn + lib.file_delimeter() + foldername
    isExist = os.path.exists(complete_output_folder)
    if not isExist:
        os.makedirs(complete_output_folder)

    driver = fdaauto.browser_init(chrome_data=args.chromedata, pdfoutput_folder=complete_output_folder)
    driver = fdaauto.browser_login(driver)
    # fdaauto.clear_screan()
    first = True
    for xlsdata in xlsdictwcode.values():
        fda_entry = FdaEntry(driver=driver, datalist=xlsdata, datearrival=args.date, pdfoutput=complete_output_folder)
        if not first:
            driver.find_element(By.CSS_SELECTOR, "img[alt='Create WebEntry Button']").click()
        
        fda_entry.parse()
        pdf_filename = fdaauto.pdf_rename(pdfoutput_folder=complete_output_folder)
        if pdf_filename != "":
            fdaauto.webentry_update(pdffile=pdf_filename, xlsfilename=destfile, pdffolder=complete_output_folder)
            xlbook.save(destfile)
        else:
            print("rename the file was failed")
        first = False
    
    list_of_files = glob.glob(complete_output_folder + lib.file_delimeter() + "*.pdf")
    allsavedfiles = []
    #regenerate data
    xlsdictall = fdaauto.xls_data_generator(xlws=xlsheet, maxrow=maxrow)
    for xlsdata in xlsdictall.values():
        entry_id = xlsdata['data'][0][20]
        pdf_filename = fdaauto.choose_pdf_file(list_of_files, entry_id)
        print('PDF File processing: ', pdf_filename)
        prior = FdaPdf(filename=pdf_filename, datalist=xlsdata, pdfoutput=complete_output_folder)
        prior.highlightpdf_generator()
        prior.insert_text()
        fdaauto.save_to_xls(pnlist=prior.pnlist)
        xlbook.save(destfile)
        allsavedfiles.extend(prior.savedfiles)
    
    setall = set(allsavedfiles)

    if len(setall) != len(allsavedfiles):
        input("Combining all pdf files Failed because there are one or more files is has the same name.")
    else:
        fdaauto.del_non_annot_page(allsavedfiles, complete_output_folder)
        fdaauto.join_folderpdf(allsavedfiles, complete_output_folder)
        # Delete all file folder
        for filename in list_of_files:
            folder = filename[:-4]
            try:
                shutil.rmtree(folder)
            except OSError as e:
                print("Error: %s : %s" % (folder, e.strerror))            
        resultfile = lib.join_pdfs(source_folder=complete_output_folder + lib.file_delimeter() + "combined", output_folder=complete_output_folder, tag="FDA_All")
        print(resultfile, "created")

        
    input("data generating completed...")


if __name__ == '__main__':
    main()
