#group names:

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font

myWorkbook = openpyxl.load_workbook("excel file name")

currentSheet = myWorkbook.active

colNum = 1
rowNum = 2

for row in currentSheet.iter_rows(min_row=rowNum, max_col=colNum, values_only=True):
    classTitle = row[0]
    if classTitle and classTitle not in myWorkbook.sheetnames:
        myWorkbook.create_sheet(title=classTitle)
myWorkbook.save("updated exel file name")