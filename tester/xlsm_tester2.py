import xlsxwriter


# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('1.xlsm')
ws = workbook.sheetnames
print(ws)