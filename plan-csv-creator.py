import xlwt
import xlrd
import csv
from shutil import copyfile
from xlutils.copy import copy
import datetime
import os
import re

def csv_from_excel(wbName,sheetName):
    wb = xlrd.open_workbook(wbName)
    sh = wb.sheet_by_name(sheetName)
    if wbName.endswith('.xlsx'):
        csvFileName = wbName[:-5]
    csvFile = open(csvFileName + '.csv', 'w')
    wr = csv.writer(csvFile, quoting=csv.QUOTE_ALL)

    for rownum in range(sh.nrows):
        wr.writerow(sh.row_values(rownum))

    csvFile.close()

def update_date_columns(wbName,sheetName):
    rb = xlrd.open_workbook(wbName)
    sh = rb.sheet_by_name(sheetName)
    row = sh.row(0)
    for colidx, cell in enumerate(row):
        if cell.value == "field_plan_availability_start" :
            startColIdx = colidx
            planStartArray = sh.col_values(colidx, 1)
        elif cell.value == "field_plan_availability_end":
            endColIdx = colidx
            planEndArray = sh.col_values(colidx, 1)
    diffStart = input('Use different start date? (y/n) ')
    if diffStart == 'y':
        startDateTimezone = input('Enter different start date (mm/dd/yyy) ')
    else:
        for start in planStartArray:
            if start != '':
                startDate = datetime.datetime(*xlrd.xldate_as_tuple(start, rb.datemode))
                startDate = startDate + datetime.timedelta(0,1)
                startDateTimezone = startDate.strftime('%m/%d/%Y %I:%M:%S')

    for end in planEndArray:
        if end != '':
            endDate = datetime.datetime(*xlrd.xldate_as_tuple(end, rb.datemode))
            endDate = endDate + datetime.timedelta(0,86399)
            endDateTimezone = endDate.strftime('%m/%d/%Y %I:%M:%S')

    wb = copy(rb)
    sheet = wb.get_sheet(0)
    for row_index in range(1, sh.nrows):
        sheet.write(row_index, startColIdx, startDateTimezone)
        sheet.write(row_index, endColIdx, endDateTimezone)    
    wb.save(wbName)

#moves all of the service areas from the zip code column to the service area column
def update_service_location_columns(wbName,sheetName):
    rb = xlrd.open_workbook(wbName)
    sh = rb.sheet_by_name(sheetName)
    row = sh.row(0)
    zipCodeList = []
    for colidx, cell in enumerate(row):
        if cell.value == "field_plan_zipcode" :
            #zipCodeColIdx = colidx
            zipCodeList = sh.col_values(colidx, 1)
    wb = copy(rb)
    sheet = wb.get_sheet(0)
    headerRow = sh.row(0)
    for colidx, cell in enumerate(headerRow):
        if cell.value == "field_plan_service_area":
            for row_index in range(1, sh.nrows):
                service_area = zipCodeList[row_index - 1].replace(', ',',')
                sheet.write(row_index, colidx, service_area)
    wb.save(wbName)

def delete_files(XWbName,copyFileName):
    del_xWb = input('Delete Excel file? (y/n) ')
    if del_xWb == 'y' and os.path.isfile(XWbName):
        os.remove(XWbName)
    if os.path.isfile(copyFileName):
        os.remove(copyFileName)

#Get the header from a template.
def get_template_header(planType):
    tenant = input('Which tenant is this for? (AK,FLB,FHCP,MN,SC,TN,WA) ')
    #if the tenant has a specific template file made use that, else use the default
    if os.path.isfile(tenant + ' - MedicareFeedTemplate.xlsx'):
        tenantTemplateName = tenant + ' - MedicareFeedTemplate.xlsx'
    else:
        tenantTemplateName = 'MedicareFeedTemplate.xlsx'
    templateCsv = xlrd.open_workbook(tenantTemplateName)
    templateSheet = templateCsv.sheet_by_name(planType)
    templateData = [templateSheet.cell_value(row, 0) for row in range(0,templateSheet.nrows)]
    return templateData

#Creates an all documents column at the end of the csv as a colmination of all of the document fields.
def create_all_documents_column(wbName,sheetName):
    rb = xlrd.open_workbook(wbName)
    sh = rb.sheet_by_name(sheetName)
    row = sh.row(0)
    regex = re.compile('field_doc_.')
    documentIdx = []
    for colidx, cell in enumerate(row):
        if re.match(regex, cell.value):
            documentIdx.append(colidx)
    print(documentIdx)

def open_excel_file():
    # get the title of the file that needs to be copied
    fileName = input("Enter file to be coverted to plan csv: ")
    if ".xlsx" not in fileName:
        fileName = fileName + '.xlsx'
    return fileName

fileName = open_excel_file()
planType = input("What plan type is to be converted? (Worksheet name) ")
copyFileName = 'Copy - ' + fileName
# copy the file in case anything happens to our original. 
copyfile(fileName,copyFileName)
openCopy = xlrd.open_workbook(copyFileName)

sheet = openCopy.sheet_by_name(planType)
# rows to skip from the top of the 
rowsToSkip = int(input("How many rows should be skipped? "))
finalWb = xlwt.Workbook()
newSheet = finalWb.add_sheet(planType)
# columns to skip when adding the data to the final csv
# will always want to skip 1 column, the labels column. if that is not there the data should not be trusted.
colToSkip = int(input("How many columns should be skipped? "))
if colToSkip == 0:
    colToSkip = 1

num_cols = sheet.ncols

for row_idx in range(rowsToSkip, sheet.nrows):
    for col_idx in range(colToSkip, num_cols):
        cell_obj = sheet.cell_value(row_idx, col_idx)
        # note: write(row,col,value) using col,row,value so that the data is in the right row form
        # removing 1 from the column to the first is blank, adding one to the top so the header is skipped
        newSheet.write(col_idx,row_idx - rowsToSkip + 1,cell_obj)

templateHeaderRowData = get_template_header(planType)
for index, value in enumerate(templateHeaderRowData):
    newSheet.write(0, index, value)

# save the final excel file
finalWbFileName = planType + ' - ' + fileName
finalWb.save(finalWbFileName)
update_date_columns(finalWbFileName, planType)
update_service_location_columns(finalWbFileName, planType)
create_all_documents_column(finalWbFileName, planType)
csv_from_excel(finalWbFileName, planType)
delete_files(finalWbFileName, copyFileName)