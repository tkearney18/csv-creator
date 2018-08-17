import xlwt
import xlrd
from shutil import copyfile

#get the title of the file that needs to be copied
fileName = input("Enter file to be coverted to plan csv: ")
if ".xlsx" not in fileName:
    fileName = fileName + '.xlsx'
planType = input("What plan type is to be converted? (Worksheet name)? ")
copyFileName = 'Copy - ' + fileName
#copy the file
copyfile(fileName,copyFileName)

finalCsv = xlrd.open_workbook(copyFileName)

sheet = finalCsv.sheet_by_name(planType)
rowsToSkip = 3
data = [sheet.cell_value(row, 0) for row in range(rowsToSkip,sheet.nrows)]
data = [''] + data
finalCsv = xlwt.Workbook()
sheet = finalCsv.add_sheet(planType)

for index, value in enumerate(data):
    sheet.write(1, index, value)

templateCsv = xlrd.open_workbook('MedicareFeedTemplate.xlsx')

templateSheet = templateCsv.sheet_by_name(planType)
templateData = [templateSheet.cell_value(row, 0) for row in range(0,templateSheet.nrows)]
for index, value in enumerate(templateData):
    sheet.write(0, index, value)
finalCsv.save(planType + ' - ' + fileName)