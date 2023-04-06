from pypdf import PdfReader
import pandas as pd
import re
import os
import argparse
import sys
from sys import platform
import time
from datetime import date
def file_delimeter():
    dm = "/" 
    if platform == "win32":
        dm = "\\"
    return dm

def get_invoice_date(fileinput):
    reader = PdfReader(fileinput)
    page = reader.pages[0]
    lines = page.extract_text().split("\n")
    return str(lines[2].split(" ")[-1])

def data_generator(fileinput):
    sourcelist = []
    reader = PdfReader(fileinput)
    number_of_pages = len(reader.pages)
    for i in range(2, number_of_pages):
        page = reader.pages[i]
        lines = page.extract_text().split("\n")
        if i == 2:
            for idx, line in enumerate(lines):
                if idx >= 12:
                    sourcelist.append(line)
        else:
            for idx, line in enumerate(lines):
                if idx >= 10:
                    sourcelist.append(line)
    return sourcelist
            
def parse_data(sourcelist):
    while True:
        if sourcelist[-1].find("Total items shipped") == -1:
            sourcelist.pop()
        else:
            sourcelist.pop()
            break
    recall = []
    date_patern = r'^\d{4}-\d{2}-\d{2}$'
    for idx, line in enumerate(sourcelist):
        cdate = line.split(" ")[0]
        if re.match(date_patern, cdate):
            rec = []
            rec.append(line)
            for idx2 in range(idx+1, len(sourcelist)):
                cdate2 = sourcelist[idx2].split(" ")[0]
                if re.match(date_patern, cdate2):
                    break
                else:
                    rec.append(sourcelist[idx2])
            recall.append(rec)
    return recall        
    
def get_result(presult):
    result = []
    regnumber = re.compile('^\d+(\.\d+)?$')
    total = 0
    for idx1, pres in enumerate(presult):
        for idx2, pr in enumerate(pres):
            if idx2 == 0:
                date = pr.split(" ")[0]
            if idx2 == 1:
                order_id = pr.split(" ")[0]

        items = []
        ituples = []
        new = True
        for idx2, pr in enumerate(pres):
            if new:
                listpr = pr.split(" ")
                if idx2 == 0:
                    ituples.append(listpr[1])
                else:
                    ituples.append(listpr[0])
                for dim in listpr:
                    if len(dim.split("x")) == 3:
                        ituples.append(dim)
                new = False
                continue
            else:
                if pr.find('Fuel Surcharge') != -1:
                    strtmp = pr.split(" ")
                    ke = 0
                    for s in strtmp:
                        if regnumber.match(s):
                            ke += 1
                            ituples.append(s)
                        if ke == 2:
                            break
            if pr[0:5] == 'Total':
                new = True
                strtotal = pr.split(" ")[1]
                total += float(strtotal[1:].strip())
                ituples.append(strtotal)
                ituples.append(float(strtotal[1:].strip()))
                items.append(tuple(ituples))
                ituples = []
                
        result.append({
            'order_date': date,
            'order_id': order_id,
            'items': items
        })

    return result

def split_values(item):
    if len(item) != 7:
        trackid = dim = billdim = weight = billweight = total = ''
        trackid = item[0]

        stext = lambda ts: [s for s in item[0:-1] if s.find(ts)!=-1]
        if len(stext('$')) == 1:
            total = stext('$')[0]

        if len(stext('x')) == 2:
            dim = stext('x')[0]
            billdim = stext('x')[1]
            weight = item[3]
            billweight = item[4]
        elif len(stext('x')) == 1:
            dim = stext('x')[0]
            weight = item[2]
            billweight = item[3]
        else:
            weight = item[1]
            billweight = item[2]
    else:
        trackid, dim, billdim, weight, billweight, total, no1 = item
    return trackid, dim, billdim, weight, billweight, total

def main():
    parser = argparse.ArgumentParser(description="Canada Post PDF Converter")
    parser.add_argument('-pdf', '--pdfinput', type=str,help="PDF File Input")
    parser.add_argument('-output', '--pdfoutput', type=str,help="PDF output folder")
    args = parser.parse_args()
        
    filelist = args.pdfinput.replace("('", '').replace("')","").replace("',)","").split("', '")
    for idx, filename in enumerate(filelist):
        isExist = os.path.exists(filename.strip())
        if not isExist:
            input(filename.strip() + " does not exist")
            sys.exit()
        else:
            filelist[idx] = filename.strip()

    isExist = os.path.exists(args.pdfoutput)
    if not isExist:
        input(args.pdfoutput + "output folder does not exist")
        sys.exit()
    print("#"*10, "Convert Canada Post PDF to CSV", "#"*10)
    reclist = []
    strdate = str(date.today())
    for file in filelist:
        basefilename = os.path.basename(file)
        print("Trying to Convert", basefilename, "...", end=" ", flush=True)
        time.sleep(2)
        try:
            sourcelist = data_generator(fileinput=file)
            invoice_date = get_invoice_date(fileinput=file)
            presult = parse_data(sourcelist=sourcelist)
            result = get_result(presult)
            # reclist = []
            for res in result:
                for item in res['items']:
                    trackid, dim, billdim, weight, billweight, total = split_values(item)
                    pdict = {
                        "Invoice Date": invoice_date,	
                        "Date": res['order_date'],	
                        "Order No":res['order_id'],
                        "Tracking": trackid,	
                        "Dimensions":dim,	
                        "Weight":weight,	
                        "Billed dimensions": billdim,
                        "Billed Weight": billweight,
                        "Total":total
                    }
                    reclist.append(pdict)

                    # pdict = {
                    #     "Invoice Date": invoice_date,	
                    #     "Date": res['order_date'],	
                    #     "Order No":res['order_id'],
                    #     "Tracking": item[0],	
                    #     "Dimensions":item[1],	
                    #     "Weight":item[3],	
                    #     "Billed dimensions": item[2],
                    #     "Billed Weight": item[4],
                    #     "Total":item[5]
                    # }
                    # reclist.append(pdict)
            # df = pd.DataFrame(reclist)
            # df.to_csv(args.pdfoutput + file_delimeter() + basefilename.replace(".pdf", "") + ".csv" , index=False)
            print("Successfully")
        except Exception as e:
            print(e, "Failed")
    df = pd.DataFrame(reclist)
    filenamesaved = "{}_{}.csv".format(args.pdfoutput + file_delimeter(), 'canada_post_pdf_convert', strdate) 
    df.to_csv(filenamesaved, index=False)
    input("End Process, Press any key to exit")
if __name__ == '__main__':
    main()