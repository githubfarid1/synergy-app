import sys
import fitz
import shutil
import os
import argparse
import time

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

class PriorPdf:
    def __init__(self, filename, page) -> None:
        print(filename, "Initialated..", end="\n\n")
        time.sleep(1)
        self.__filename = filename
        self.__page = page
        try:
            doc = fitz.open(self.filename)
            page = doc[page]
            self.__word_list = page.get_text("words")
        except:
            input("file not found")
            sys.exit()

    def generate(self):
        codelist = []
        print("Process Starting...")
        time.sleep(1)
        for word in self.word_list:
            if word[0] == POSX1CODE1 and word[2] == POSX2CODE1:
                dict = {'key': word[4], 'value': word}
                codelist.append(dict)
        folder = os.path.dirname(self.filename)
        tmpfile = "{}/temp.pdf".format(folder)

        for key, code in enumerate(codelist):
            newfilename = "{}/{}.pdf".format(folder, code['key'])
            shutil.copyfile(self.filename, tmpfile)
            doc = fitz.open(tmpfile)
            page = doc[self.page]
            r = fitz.Rect(code['value'][0], code['value'][1], code['value'][2], code['value'][3])
            page.add_highlight_annot(r)
            try:
                r = fitz.Rect(POSX1DESC, code['value'][1], POSX2DESC, codelist[key+1]['value'][3] - 13)
                page.add_highlight_annot(r)
            except:
                r = fitz.Rect(POSX1DESC, code['value'][1], POSX2DESC, code['value'][3] + 17)
                page.add_highlight_annot(r)
            
            r = fitz.Rect(POSX1LOC, code['value'][1], POSX2LOC, code['value'][3])
            page.add_highlight_annot(r)
            r = fitz.Rect(POSX1DATE, code['value'][1], POSX2DATE, code['value'][3])
            page.add_highlight_annot(r)
            r = fitz.Rect(POSX1CODE2, code['value'][1], POSX2CODE2, code['value'][3]+2)
            page.add_highlight_annot(r)
            doc.save(newfilename)
            print(newfilename, "Created")
            time.sleep(1)
        doc.close()
        os.remove(tmpfile)
        input("Generating PDF File Completed..")
    @property
    def filename(self):
        return self.__filename


    @filename.setter
    def filename(self, value):
        self.__filename.set(value)

    @property
    def page(self):
        return self.__page


    @page.setter
    def page(self, value):
        self.__page.set(value)

    @property
    def word_list(self):
        return self.__word_list


def main():
    parser = argparse.ArgumentParser(description="FDA PDF Extractor")
    parser.add_argument('-pdf', '--input', type=str,help="PDF File Input")
    args = parser.parse_args()
    if args.input[-4:] != '.pdf':
        print('1st file input have to PDF file')
        sys.exit()

    prior = PriorPdf(args.pdfinput, 0)
    prior.generate()
    
if __name__ == '__main__':
    main()


  