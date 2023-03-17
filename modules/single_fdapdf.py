import sys
import fitz
import shutil
import os
import time
from sys import platform
from os import walk

POSX1CODE1 = 12.0
POSX2CODE1 = 32.01737594604492
POSX1DESC = 71.0999984741211
POSX2DESC = 271.55694580078125
POSX1LOC = 277.95001220703125
POSX2LOC = 290.450927734375
POSX1DATE = 396.1499938964844
POSX2DATE = 478.6723937988281
POSX1CODE2 = 514.3499755859375
POSX2CODE2 = 594.415771484375

def clearlist(*args):
    for varlist in args:
        varlist.clear()

def combine_allpdf(pdffiles, pdfoutput):
    print("Merging PDF files started..")
    file_delimeter = "/"    
    if platform == "win32":
        file_delimeter = "\\"

    time.sleep(1)
    dictfiles = {}
    result = fitz.open()
    for pdffile in pdffiles:
        basefilename = os.path.basename(pdffile)
        dictfiles[int(basefilename.replace(".pdf",""))] = pdffile
    sortedfiles = dict(sorted(dictfiles.items()))
    # print(sortedfiles)

    for file in sortedfiles:
        print(os.path.basename(sortedfiles[file]), "merged")
        mfile = fitz.open(sortedfiles[file])
        result.insert_pdf(mfile)
        time.sleep(1)
    result.save(pdfoutput + file_delimeter + "combine.pdf")
    print(pdfoutput + file_delimeter + "combine.pdf", "Created!")
    print("Merging PDF files finished..")

class FdaPdf:
    def __init__(self, filename, datalist, pdfoutput) -> None:
        print(os.path.basename(filename), "Initialated FDA PDF..")
        time.sleep(1)
        self.__filename = filename
        self.__pnlist = []
        self.__pdfoutput = pdfoutput
        self.__savedfiles = []
        self.__datalist = datalist
        self.__delimeter = "/"    
        if platform == "win32":
            self.__delimeter = "\\"

        self.__newpdf_generator()
    
    def __newpdf_generator(self):
        foldername = self.filename.replace(".pdf",'')
        isExist = os.path.exists(foldername)
        if not isExist:
            os.makedirs(foldername)
        boxlist = []
        for ds in self.datalist['data']:
            tmpboxes = ds[19].replace('Box','').split(',')
            for tmpbox in tmpboxes:
                fname = "{}.pdf".format(tmpbox.strip())
                boxlist.append(fname)

        for box in set(boxlist):
            shutil.copy(self.filename, "{}{}{}".format(foldername, self.file_delimeter, box) )
        
    def insert_text(self):
        print("")
        print("Inserting filenames text into PDF started..")
        time.sleep(1)
        pdffolder = self.filename.replace(".pdf",'')
        filenames = next(walk(pdffolder), (None, None, []))[2]  # [] if no file
        if platform == "linux" or platform == "linux2":
            pdffolder = pdffolder + "/"
        elif platform == "win32":
            pdffolder = pdffolder + "\\"
        red = fitz.utils.getColor("red")
        for filename in filenames:
            shutil.copy(pdffolder + self.file_delimeter + filename, pdffolder + self.file_delimeter + "tmp.pdf")
            doc = fitz.open(pdffolder + self.file_delimeter + "tmp.pdf")
            for i in range(0, doc.page_count):
                page = doc[i]
                page.insert_text((520.2469787597656, 803.38037109375), filename, color=red)
            doc.save(pdffolder + self.file_delimeter + filename)
            self.savedfiles.append(pdffolder + self.file_delimeter + filename)
            print(filename, "inserted.")
            time.sleep(1)
            doc.close()    
        os.remove(pdffolder + self.file_delimeter + "tmp.pdf")
        print("Inserting the filenames finished..", end="\n----------------------------------------\n\n")
    
    def __research_text(self, pdfpage, text):
        for i in range(0, len(text)+1):
            rect = pdfpage.search_for(text[0:i],flags=(fitz.TEXT_PRESERVE_WHITESPACE))
            if rect == []:
                break
        lastfound = text[0:i-1]
        tail = text.replace(lastfound, "")
        textsearch = lastfound + " " + tail
        return pdfpage.search_for(textsearch,flags=(fitz.TEXT_PRESERVE_WHITESPACE))
        # pass
    def highlightpdf_generator(self):
        foldername = self.filename.replace(".pdf",'')
        print("")
        print("Highlight text Process Starting...")
        time.sleep(1)
        code2set = set()
        for ds in self.datalist['data']:
            time.sleep(1)
            print(ds[20], ds[19])
            tmpboxes = ds[19].replace('Box','').split(',')
            for tmpbox in tmpboxes:
                fname = "{}.pdf".format(tmpbox.strip())
                shutil.copy(foldername + self.file_delimeter + fname, foldername + self.file_delimeter + "tmp.pdf")
                doc = fitz.open(foldername + self.file_delimeter + "tmp.pdf")
                for i in range(0, doc.page_count):
                    pdfpage = doc[i]
                    rects = pdfpage.search_for(ds[2], flags=(fitz.TEXT_PRESERVE_WHITESPACE))
                    
                    if rects == []:
                        rects = self.__research_text(pdfpage, ds[2])
                        if rects != []:
                            break
                    else:
                        break
                if rects == []:
                    input("Item not found, Report to administrator")
                    sys.exit()
                
                splitter = rects[0][2]
                #DON'T REMOVE THIS
                # try:
                #     splitter = rects[0][2]
                # except:
                #     words = ds['description'].split()
                #     res = ''
                #     for i in range(1, len(words)):
                #         rects = pdfpage.search_for(' '.join(words[:i]))
                #         if rects == []:
                            
                #             res = ' '.join(words[:i-1])
                #             break
                #     rects = pdfpage.search_for(res)
                #     rects[0][3] = rects[0][3] + 10
                #     splitter = rects[0][2]


                # try:
                #     splitter = rects[0][2]
                #     # print(ds['description'])
                # except:
                #     words = ds['description'].split()
                #     pos = -1
                #     while True:
                #         sent = ' '.join(words[:pos])
                #         try:
                #             # print(sent)
                #             rects = pdfpage.search_for(sent)
                #             print(rects)
                #             splitter = rects[0][2]
                #             break
                #         except:
                #             pos = pos-1
                #             continue
                # print(rects)            
                
                rdata = []
                tmpdata = []
            
                first = True
                for rect in rects:
                    # print(rect[2])
                    if first:
                        tmpdata.append(rect)
                        first = False
                    else:    
                        if rect[2] == splitter:
                            rdata.append(tmpdata)
                            tmpdata = []
                            tmpdata.append(rect)
                            # pass
                        else:
                            tmpdata.append(rect)
                
                rdata.append(tmpdata)

                if len(rdata) > 1:
                    for rd in rdata:
                        pncode1s = pdfpage.get_text("blocks", clip=(POSX1CODE1, rd[0][1]-10, POSX2CODE1, rd[0][3]+10))
                        pncode2s = pdfpage.get_text("blocks", clip=(POSX1CODE2, rd[0][1]-10, POSX2CODE2, rd[0][3]+10))
                        locs = pdfpage.get_text("blocks", clip=(POSX1LOC, rd[0][1]-10, POSX2LOC, rd[0][3]+10))
                        dates = pdfpage.get_text("blocks", clip=(POSX1DATE, rd[0][1]-10, POSX2DATE, rd[0][3]+10))
                        # print(pncode2s[0][4])
                        if pncode2s[0][4].strip() in code2set:
                            continue
                        else:
                            code2set.add(pncode2s[0][4].strip())
                            pnnumber = pncode2s[0][4].strip()
                            r = fitz.Rect(pncode2s[0][0], pncode2s[0][1], pncode2s[0][2], pncode2s[0][3])
                            pdfpage.add_highlight_annot(r)
                            r = fitz.Rect(pncode1s[0][0], pncode1s[0][1], pncode1s[0][2], pncode1s[0][3])
                            pdfpage.add_highlight_annot(r)
                            r = fitz.Rect(locs[0][0], locs[0][1], locs[0][2], locs[0][3])
                            pdfpage.add_highlight_annot(r)
                            r = fitz.Rect(dates[0][0], dates[0][1], dates[0][2], dates[0][3])
                            pdfpage.add_highlight_annot(r)
                            for rx in rd:
                                r = fitz.Rect(rx)
                                pdfpage.add_highlight_annot(r)

                            doc.save(foldername + self.file_delimeter + fname)

                            break
                else:
                    pncode1s = pdfpage.get_text("blocks", clip=(POSX1CODE1, rdata[0][0][1]-10, POSX2CODE1, rdata[0][0][3]+10))
                    pncode2s = pdfpage.get_text("blocks", clip=(POSX1CODE2, rdata[0][0][1]-10, POSX2CODE2, rdata[0][0][3]+10))
                    locs = pdfpage.get_text("blocks", clip=(POSX1LOC, rdata[0][0][1]-10, POSX2LOC, rdata[0][0][3]+10))
                    dates = pdfpage.get_text("blocks", clip=(POSX1DATE, rdata[0][0][1]-10, POSX2DATE, rdata[0][0][3]+10))
                    pnnumber = pncode2s[0][4].strip()
                    r = fitz.Rect(pncode2s[0][0], pncode2s[0][1], pncode2s[0][2], pncode2s[0][3])
                    pdfpage.add_highlight_annot(r)
                    r = fitz.Rect(pncode1s[0][0], pncode1s[0][1], pncode1s[0][2], pncode1s[0][3])
                    pdfpage.add_highlight_annot(r)
                    r = fitz.Rect(locs[0][0], locs[0][1], locs[0][2], locs[0][3])
                    pdfpage.add_highlight_annot(r)
                    r = fitz.Rect(dates[0][0], dates[0][1], dates[0][2], dates[0][3])
                    pdfpage.add_highlight_annot(r)

                    for rx in rdata[0]:
                        r = fitz.Rect(rx)
                        pdfpage.add_highlight_annot(r)

                    doc.save(foldername + self.file_delimeter + fname)
                    #pnlist
                # self.datalist['data'][idx][20] = pnnumber
            dict = {'entry_id': ds[20], 'pnnumber': pnnumber, 'boxes': ds[19], 'sku': ds[21]}
            self.pnlist.append(dict)
        try:
            doc.close()
            os.remove(foldername + self.file_delimeter + "tmp.pdf")
        except:
            input("No Data found, make sure you have input `Web Entry Identifier` in the Excel file..")
            sys.exit()
    
    def combine_allpdf(self, pdffiles):
        print("Merging PDF files started..")
        time.sleep(1)
        dictfiles = {}
        result = fitz.open()
        for pdffile in pdffiles:
            basefilename = os.path.basename(pdffile)
            dictfiles[int(basefilename.replace(".pdf",""))] = pdffile
        sortedfiles = dict(sorted(dictfiles.items()))
        # print(sortedfiles)

        for file in sortedfiles:
            print(os.path.basename(sortedfiles[file]), "merged")
            mfile = fitz.open(sortedfiles[file])
            result.insert_pdf(mfile)
            time.sleep(1)
        result.save(self.pdfoutput_folder + self.file_delimeter + "combine.pdf")
        print(self.pdfoutput_folder + self.file_delimeter + "combine.pdf", "Created!")
        print("Merging PDF files finished..")

    @property
    def filename(self):
        return self.__filename


    @filename.setter
    def filename(self, value):
        self.__filename.set(value)

    @property
    def datalist(self):
        return self.__datalist

    @property
    def pnlist(self):
        return self.__pnlist

    @property
    def savedfiles(self):
        return self.__savedfiles

    @property
    def pdfoutput_folder(self):    
        return self.__pdfoutput

    @property
    def file_delimeter(self):
        return self.__delimeter


if __name__ == '__main__':
    print('This module can not run via Main')
