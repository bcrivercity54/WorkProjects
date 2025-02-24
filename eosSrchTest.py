import re
import os
import fnmatch
import openpyxl
from openpyxl.cell import MergedCell



techList= []


def getTech(xFile):

    wb = openpyxl.load_workbook(xFile, read_only=False)
    sheet_ranges = wb['Sheet1']
    ws = wb.active
    for row in ws.iter_rows(2):
        for cell in row:
            print("cell value",cell.value)
            if cell.value == "Tech:" and cell.offset(row=0, column=1).value is not  None:
                #print("cell value",cell.value)
                #rowNum ="row("+row+")"
                rowNum = str((row[0:1]))
                sRowNum = rowNum.replace("<Cell 'Sheet1'.",'')
                sRowNum = sRowNum.replace(">,",'')
                print("Row num",sRowNum)
                #print(row,cell.value,cell.offset(row=0, column=1).value)
                #outStringList.append(sRowNum) #label.config(text=sRowNum, font=('Courier 13 bold'))

    


rootPath = r'C:\Users\brian\Documents\Spreadsheets\Work_Instructions'

custtext = "GE_HC_OEM"

refText = "4283348"

#pattern = re.search("\\d{7}", txt)


searchPath = os.path.join(rootPath, custtext)
#print("sp",searchPath)
#m = re.search(refText,searchPath)
#print("first m",m)
#print("searchPath",searchPath)


for filename in os.listdir(searchPath):
    #print(filename)
    searchfile = os.path.join(searchPath,filename)
    m = re.search("\\d{7}",searchfile)
    if m.group() == refText:
        print("awesome",searchfile)
        techList.append(searchfile)
    #print("second m",m.group(),searchfile)
    #print('sf',searchfile)
    if fnmatch.fnmatch(searchfile,"4283348"):
        print("true")

for item in techList:
    getTech(item)

    
    #print("tech find",item)
