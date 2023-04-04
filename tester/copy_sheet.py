#tesss
import sys
sys.path.insert(0, '../modules')

import amazon_lib as lib
# from openpyxl import Workbook, load_workbook

# filename1 = "/home/farid/dev/python/synergy-github/data/sample/copy_sheet/xUSA Small Shipment Creation V12.20.xlsm"
# filename2 = "/home/farid/dev/python/synergy-github/data/sample/copy_sheet/April 01 Labels.xlsx"
destination = r"C:/synergy-data-tester/shipmentall/xUSA Small Shipment Creation V12.20.xlsm"
source = r"C:/synergy-data-tester/shipmentall/shipment_creation_2023-04-02/April 02 Labels.xlsx"
cols = ('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H')
sheetsource="Sheet"
sheetdestination="Shipment labels summary"
tracksheet= "dyk_manifest_template"

lib.copysheet(destination=destination, source=source, cols=('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'), sheetsource="Sheet", sheetdestination="Shipment labels summary", tracksheet="dyk_manifest_template")
