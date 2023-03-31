import amazon_lib as lib
import argparse
import sys
import os
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(description="Amazon Shipment Join")
    parser.add_argument('-sourcefolder', '--sourcefolder', type=str,help="Source PDF Folder")
    parser.add_argument('-outputfolder', '--outputfolder', type=str,help="Output PDF Folder")
    args = parser.parse_args()
    isExist = os.path.exists(args.sourcefolder)
    if not isExist:
        input(args.sourcefolder + " does not exist")
        sys.exit()

    isExist = os.path.exists(args.outputfolder)
    if not isExist:
        input(args.outputfolder + " does not exist")
        sys.exit()

    addressfile = Path("address.csv")
    resultfile = lib.join_pdfs(source_folder=args.sourcefolder, output_folder = args.outputfolder, tag='Labels')
    if resultfile != "":
        lib.add_page_numbers(resultfile)
        lib.generate_xls_from_pdf(resultfile, addressfile)
    input("End Process..")    

if __name__ == '__main__':
    main()
