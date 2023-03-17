import fitz
import os
from openpyxl import Workbook, load_workbook
import sys

def extract_pdf(box, shipment_id, label):
    pdffile = "{}{}package-{}.pdf".format('/home/farid/dev/python/synergy-github/data/sample/tesfeb24', "/", shipment_id)
    foldername = "{}{}combined".format('/home/farid/dev/python/synergy-github/data/sample/tesfeb24', "/") 
    isExist = os.path.exists(foldername)
    if not isExist:
        os.makedirs(foldername)
    # boxes = dlist
    white = fitz.utils.getColor("white")

    # print(box, pdffile)
    mfile = fitz.open(pdffile)
    fname = "{}{}{}.pdf".format(foldername, "/",  box.strip())
    tmpname = "{}{}{}.pdf".format(foldername, "/", "tmp")
    isExist = os.path.exists(tmpname)
    if isExist:
        os.remove(tmpname)
    # input("pause")    
    found = False
    pfound = 0
    for i in range(0, mfile.page_count):
        page = mfile[i]
        plist = page.search_for(label)
        if len(plist) != 0:
            found = True
            pfound = i
            print(i)
            break
    if found:
        single = fitz.open()
        single.insert_pdf(mfile, from_page=pfound, to_page=pfound)
        mfile.close()
        single.save(tmpname)
        mfile = fitz.open(tmpname)
        page = mfile[0]
        page.insert_text((550.2469787597656, 100.38037109375), "Box:{}".format(str(box)), rotate=90, color=white)
        page.set_rotation(90)
        mfile.save(fname)




workbook = load_workbook(filename='/home/farid/dev/python/synergy-github/data/sample/Shipment transfer.xlsx', read_only=False, data_only=True)
worksheet = workbook['Feb24']


shiplist = []
dtmp = []
dict = {
    'shipmentid': 'FBA1717VBSZH',
    'label':'FBA1717VBSZHU000001',
    'trackid': '1ZR726A70395513664',
    'weight': 23,
    'dimension': "13 x 13 x 13",

    }
dtmp.append(dict)
shiplist.append(dtmp)
dtmp = []
dict = {
    'shipmentid': 'FBA1717TSLT3',
    'label':'FBA1717TSLT3U000001',
    'trackid': '1ZR726A70350983679',
    'weight': 22,
    'dimension': '12 x 12 x 12',

    }    
dtmp.append(dict)
shiplist.append(dtmp)
# print(shiplist)
boxes = ('E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P')
dlist =  {
    "begin": 95,
    "end": 117,
    "submitter": "DMT Distributors (102 Central Avenue)",
    "address": "CLT2 (746 NW 7th Ave Florida City,  FL  33034)",
    "name": "Shipment 5",
    "boxcount": 2,
    "weightboxes": [
      22,
      23
    ],
    "dimensionboxes": [
      "12x12x12",
      "13x13x13"
    ],
    "nameboxes": [
      "1",
      "3"
    ],
}

# shiplist = []
# dtmp = []
# dict = {
#     'shipmentid': 'FBA1717TPZYN',
#     'label':'FBA1717TPZYNU000001',
#     'trackid': '1ZR726A70331965680',
#     'weight': 21,
#     'dimension': "12 x 12 x 12",
#     }
# dtmp.append(dict)
# dict = {
#     'shipmentid': 'FBA1717TPZYN',
#     'label':'FBA1717TPZYNU000002',
#     'trackid': '1ZR726A70325499295',
#     'weight': 37,
#     'dimension': "14 x 14 x 14",
#     }
# dtmp.append(dict)

# shiplist.append(dtmp)
# # print(shiplist)
# boxes = ('E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P')
# dlist =  {
#     "begin": 3,
#     "end": 25,
#     "submitter": "DMT Distributors (102 Central Avenue)",
#     "address": "SCK4 (102 Central Ave Ste 6540)",
#     "name": "Shipment 1",
#     "boxcount": 2,
#     "weightboxes": [
#       21,
#       37
#     ],
#     "dimensionboxes": [
#       "12x12x12",
#       "14x14x14"
#     ],
#     "nameboxes": [
#       "2",
#       "7"
#     ],
# }

boxcols = ('E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P')
for ship in shiplist:
    for s in ship:
        for boxcol in boxcols:
            dimension = ""
            weight = ""
            box = str(worksheet['{}{}'.format(boxcol, dlist['begin'])].value)
            # print(box)
            if box != 'None':
                for i in range(dlist['begin'], dlist['end']):
                    if worksheet['B{}'.format(i)].value == 'Weight':
                        weight = worksheet['{}{}'.format(boxcol, i)].value
                    
                    if worksheet['B{}'.format(i)].value == 'Dimensions':
                        # print("xxx")
                        dimension = worksheet['{}{}'.format(boxcol, i)].value

                    dimension = dimension.replace(" ","")
                    dimship = s['dimension'].replace(" ","")

                if s['weight'] == int(weight) and dimension == dimship:
                    print(weight, dimension, box)
                    extract_pdf(box=box, shipment_id=s['shipmentid'], label=s['label'])

