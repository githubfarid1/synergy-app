import xlwings as xw
wb = xw.Book('x.xlsm')
sheet = wb.sheets['Shipment summary']
sheet.range('A1').value = 'From Script'
wb.save('result_file_name.xlsm')