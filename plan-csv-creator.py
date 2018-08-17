import xlwt
import xlrd
import csv
from shutil import copyfile

def csv_from_excel(wbName,sheetName):
    wb = xlrd.open_workbook(wbName)
    sh = wb.sheet_by_name(sheetName)
    your_csv_file = open('your_csv_file.csv', 'w')
    wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)

    for rownum in range(sh.nrows):
        wr.writerow(sh.row_values(rownum))

    your_csv_file.close()

# get the title of the file that needs to be copied
fileName = input("Enter file to be coverted to plan csv: ")
if ".xlsx" not in fileName:
    fileName = fileName + '.xlsx'
planType = input("What plan type is to be converted? (Worksheet name) ")
copyFileName = 'Copy - ' + fileName
# copy the file in case anything happens to our original. 
copyfile(fileName,copyFileName)
openCopy = xlrd.open_workbook(copyFileName)

sheet = openCopy.sheet_by_name(planType)
# rows to skip from the top of the 
rowsToSkip = 3
# columns to skip when adding the data to the final csv
columnsToSkip = 1
finalWb = xlwt.Workbook()
newSheet = finalWb.add_sheet(planType)
num_cols = sheet.ncols - 2
for row_idx in range(rowsToSkip, sheet.nrows):
    for col_idx in range(columnsToSkip + 1, num_cols):
        cell_obj = sheet.cell_value(row_idx, col_idx)
        newSheet.write(col_idx - 1,row_idx - rowsToSkip + 1,cell_obj)

templateCsv = xlrd.open_workbook('MedicareFeedTemplate.xlsx')

templateSheet = templateCsv.sheet_by_name(planType)
templateData = [templateSheet.cell_value(row, 0) for row in range(0,templateSheet.nrows)]
for index, value in enumerate(templateData):
    newSheet.write(0, index, value)
# save the final excel file
finalWbFileName = planType + ' - ' + fileName
finalWb.save(finalWbFileName)
csv_from_excel(finalWbFileName, planType)