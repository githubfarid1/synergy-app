# from openpyxl import Workbook, load_workbook

# xlsfile = 'C:\\synergy-data-tester\\shipmentall\\xUSA Small Shipment Creation V12.20.xlsm'
# sname = 'Shipment summary'
# workbook = load_workbook(filename=xlsfile, read_only=False, keep_vba=True, data_only=True)
# worksheet = workbook[sname]
# workbook.save(xlsfile)

import xlwings as xw
xlsfile = r'C:/synergy-data-tester/shipmentall/xUSA Small Shipment Creation V12.20.xlsm'
newfile = r'C:/synergy-data-tester/shipmentall/yUSA Small Shipment Creation V12.20.xlsm'
sname = 'Shipment summary'
sheet1 = xw.Book(xlsfile).sheets[sname]
sheet1['A4'].value = 'Hellox'
xw.Book(xlsfile).save()
xw.Book(xlsfile).close()