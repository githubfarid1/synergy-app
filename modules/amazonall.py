import sys
import os
import argparse
import time
from sys import platform
import json
from random import randint
from datetime import date, datetime, timedelta
import amazon_lib as lib
from subprocess import Popen

if platform == "linux" or platform == "linux2":
    PYLOC = "python"
elif platform == "win32":
    PYLOC = "python.exe"

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
    foldername = "{}{}_{}".format(args.pdfoutput + lib.file_delimeter(), 'shipment_creation', strdate) 
    isExist = os.path.exists(foldername)
    if not isExist:
        os.makedirs(foldername)
    print("Step 1: Shipment Creation")
    comlist=[PYLOC, "modules/amazonship.py", "-xls", args.xlsinput, "-sname", args.shipsheet, "-output", foldername, "-cdata",  args.chromedata]
    Popen(comlist)
    foldername = "{}{}_{}".format(args.pdfoutput + lib.file_delimeter(), 'prior_notice', strdate) 
    isExist = os.path.exists(foldername)
    if not isExist:
        os.makedirs(foldername)

    print("Step 2: Prior Notice")
    comlist=[PYLOC, "modules/autofdapdf.py", "-i", args.xlsinput, "-d", args.chromedata, "-s", args.shipsheet, "-dt", args.date, "-o", foldername]
    Popen(comlist)


if __name__ == '__main__':
    main()
