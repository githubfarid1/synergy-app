from fdaentry import FdaEntry
from fdapdf import PriorPdf
import argparse
import sys
from sys import platform
import os
import shutil
import time
import fitz
def clear_screan():
    if platform == "win32":
        os.system("cls")
    else:
        os.system("clear")
def file_delimeter():
    delimeter = "/"    
    if platform == "win32":
        delimeter = "\\"
    return delimeter

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
        time.sleep(1)
    isExist = os.path.exists(tmpfile)
    doc.close()
    if isExist:    
        os.remove(tmpfile)    
    print("")

def join_folderpdf(pdffiles, pdfoutput_folder):
    print("Merging PDF files in one folder started..")
    time.sleep(1)

    foldername = pdfoutput_folder + file_delimeter() + "combined"
    isExist = os.path.exists(foldername)
    if isExist:
        try:
            shutil.rmtree(foldername)
        except OSError as e:
            print("Error: %s : %s" % (foldername, e.strerror))            
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


def main():
    parser = argparse.ArgumentParser(description="FDA Entry + PDF Extractor")
    parser.add_argument('-i', '--input', type=str,help="File Input")
    parser.add_argument('-s', '--sheet', type=str,help="Sheet Name")
    parser.add_argument('-dt', '--date', type=str,help="Arrival Date")
    parser.add_argument('-d', '--chromedata', type=str,help="Chrome User Data Directory")
    parser.add_argument('-o', '--output', type=str,help="PDF output folder")
    
    args = parser.parse_args()
    if args.input[-5:] != '.xlsx':
        input('File input have to XLSX file')
        sys.exit()
    isExist = os.path.exists(args.chromedata)
    if isExist == False :
        input('Please check Chrome User Data Directory')
        sys.exit()
    isExist = os.path.exists(args.input)
    if isExist == False :
        input('Please check XLSX file')
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
    fdaentry = FdaEntry(args.input, args.sheet, args.date, args.output, args.chromedata)
    fdaentry.xls_data_generator()
    fdaentry.parse()
    fdaentry.pdf_rename()
    fdaentry.webentry_update()
    # sys.exit()
    allsavedfiles = []
    for filename in fdaentry.pdffilelist:
        prior = PriorPdf(filename, 0, args.input, args.sheet, args.output)
        prior.highlightpdf_generator()
        prior.save_to_xls()
        prior.insert_text()
        allsavedfiles.extend(prior.savedfiles)

    setall = set(allsavedfiles)
    if len(setall) != len(allsavedfiles):
        input("Combining all pdf files Failed because there are one or more files is has the same name.")
    else:
        del_non_annot_page(allsavedfiles, args.output)
        join_folderpdf(allsavedfiles, args.output)
        # Delete all file folder
        for filename in fdaentry.pdffilelist:
            folder = filename[:-4]
            try:
                shutil.rmtree(folder)
            except OSError as e:
                print("Error: %s : %s" % (folder, e.strerror))            
            
    input("data generating completed...")

if __name__ == '__main__':
    main()
