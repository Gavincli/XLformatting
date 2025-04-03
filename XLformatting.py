# Mary Catherine Shepherd, Sam Jenson, Gavin Clifton, Ben Funk, Thomas Apke
# professor anderson, section 003


import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font

#insert path to the file!!
myWorkbook = openpyxl.load_workbook()
# setting up the worksheet to be the active sheet in our file
currentSheet = myWorkbook.active


# setting up variables
colNum = 1

rowNum = 2
# creating a for loop to split up names and ID, and to create new sheets
for row in currentSheet.iter_rows(min_row=rowNum, values_only=True):
    classTitle = row[0]
    lstStudInfo = row[1].split('_')
    lstStudInfo.append(row[2])

    if classTitle not in myWorkbook.sheetnames:
        printRow = 1
        myWorkbook.create_sheet(title=classTitle)
        myWorkbook[classTitle]['A1'] = 'First Name'
        myWorkbook[classTitle]['B1'] = 'Last Name'
        myWorkbook[classTitle]['C1'] = 'Student ID'
        myWorkbook[classTitle]['D1'] = 'Grade'

    #move down a row per iteration
    printRow += 1
    #print values on the new sheet as long as the classTitle matches
    #myWorkbook[classTitle]['A' + str(printRow)].append(lstStudInfo)
    myWorkbook[classTitle]['A' + str(printRow)] = lstStudInfo[0]
    myWorkbook[classTitle]['B' + str(printRow)] = lstStudInfo[1]
    myWorkbook[classTitle]['C' + str(printRow)] = lstStudInfo[2]
    myWorkbook[classTitle]['D' + str(printRow)] = lstStudInfo[3]

        worksheet = myWorkbook.create_sheet(title=classTitle)



#apply filters for each worksheet
for worksheet in myWorkbook.worksheets:
    max_row = worksheet.max_row
    worksheet.auto_filter.ref = f"A1:D{max_row}"

# formats each column with bold font and proper size
for col in ['A', 'B', 'C', 'D', 'F', 'G']:
    sheet[f'{col}1'].font = Font(bold = True)
    sheet.column_dimensions[col].width = len(sheet[f'{col}1'].value) + 5

# Saves the new file
myWorkbook.save('formatted_grades.xlsx')
