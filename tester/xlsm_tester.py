from openpyxl import Workbook, load_workbook

xlsfile = '/home/farid/dev/python/synergy-github/data/sample/xUSA Small Shipment Creation V12.20.xlsm'
sname = 'Shipment summary'
workbook = load_workbook(filename=xlsfile, read_only=False, keep_vba=True, data_only=True)
worksheet = workbook[sname]
