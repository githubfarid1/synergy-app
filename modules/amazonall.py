import sys
import os
import argparse
import time
from sys import platform
import json
from random import randint
from datetime import date, datetime, timedelta
import amazon_lib as lib
import amazonship
import logging
from pathlib import Path

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

    logger2.info("###### Start ######")
    logger2.info("Filename: {}\nSheet Name:{}\nPDF Output Folder:{}".format(args.xlsinput, args.shipsheet, args.pdfoutput))
    maxrun = 10
    for i in range(1, maxrun+1):
        if i > 1:
            print("Process will be reapeated")
        try:    
            shipment = amazonship.AmazonShipment(xlsfile=args.xlsinput, sname=args.shipsheet, chrome_data=args.chromedata, download_folder=args.pdfoutput)
            shipment.data_sanitizer()
            if len(shipment.datalist) == 0:
                break
            shipment.parse()
        except Exception as e:
            logger.error(e)
            print("There is an error, check logs/amazonship-err.log")
            shipment.workbook.save(shipment.xlsfile)
            shipment.workbook.close()
            if i == maxrun:
                logger.error("Execution Limit reached, Please check the script")
            continue
        break
    addressfile = Path("address.csv")
    resultfile = lib.join_pdfs(source_folder=args.pdfoutput + lib.file_delimeter() + "combined" , output_folder = args.pdfoutput, tag='Labels')
    if resultfile != "":
        lib.add_page_numbers(resultfile)
        lib.generate_xls_from_pdf(resultfile, addressfile)
    input("End Process..")    



if __name__ == '__main__':
    main()
