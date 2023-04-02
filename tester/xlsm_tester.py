# from openpyxl import Workbook, load_workbook

# xlsfile = 'C:\\synergy-data-tester\\shipmentall\\xUSA Small Shipment Creation V12.20.xlsm'
# sname = 'Shipment summary'
# workbook = load_workbook(filename=xlsfile, read_only=False, keep_vba=True, data_only=True)
# worksheet = workbook[sname]
# workbook.save(xlsfile)

from xlwings import Workbook
xlsfile = r'C:/synergy-data-tester/shipmentall/xUSA Small Shipment Creation V12.20.xlsm'
newfile = r'C:/synergy-data-tester/shipmentall/yUSA Small Shipment Creation V12.20.xlsm'
sname = 'Shipment summary'
wb = Workbook(xlsfile)
wb.save(newfile)
