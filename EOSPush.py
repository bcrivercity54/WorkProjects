import sys
import os
from PyQt5.uic import loadUi
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QDialog, QApplication
from pathlib import Path
import glob
import fnmatch
from os import listdir
from os.path import isfile, join

testDict = {}

class EOS(QDialog):
    def __init__(self):
        super(EOS, self).__init__()
        loadUi("test3.ui", self)
        self.tableWidget.setColumnWidth(0, 250)
        self.tableWidget.setColumnWidth(1, 250)
        self.tableWidget.setColumnWidth(2, 250)

        # print('test')

        self.grabButton.clicked.connect(self.getbwi)

        #self.loaddata()

    #def loaddata(self):




    def getbwi(self=None):
        print("in the func")
        custtext = self.textCustomer.toPlainText()
        refText = self.refNum.toPlainText()
        #print("refText",refText)

        print(custtext,refText)
        #joinRootToCustomer(custtext)

        rootPath = r'C:\Users\brian\Documents\Spreadsheets\Work_Instructions'

        searchPath = os.path.join(rootPath, custtext)
        print("search path",searchPath)

        pattern = refText
        #print("pattern",type(refText))
        # Search the folder for the entered ref number add to a list or dictionary

        fileList = []

        #for root, dirs, files in os.walk(searchPath):
            #print(files,pattern)
            #if fnmatch.filter(files,'*pattern*'):


        #f fnmatch.filter(i,"*_4283348_*.xlsx"):
                # fileList.append(i)
                #print("Found files",files)
        #for root, dirs, files in os.walk(searchPath):
           #print(searchPath)
            #if fnmatch.filter(files, "*_refText_*.xlsx"):
                #print("if",files)
        # fileList.append(i)
            #for filename in fnmatch.filter(files,'*_refText_*.xlsx'):
                #fileList.append(filename)
                #print("FUDGE",filename)


        #for x in fileList:
            #print(x)
            #print("End of processing")
        #for i in os.listdir(searchPath):
           #print("searchpath",searchPath)

            #if fnmatch.filter(i,"*_refText_*.xlsx"):
                #fileList.append(i)

                #print(i)
        fileList = ["GE_MR_OEM_4242234_9X3M2Y.xlsm", "GE_MR_OEM_4242234_8X3M2Y.xlsm", "GE_MR_OEM_4242234_7X3M2Y.xlsm"]

        customerInfoList = []
        refNumList = []
        serviceTagList = []
        techName = ['Brian', 'Brian', 'Brian']
        r = 0
        for item in fileList:
            custtDict = {"cust":item[-35:-20]}
            #,"customer":item[-11:-5],"servicetag":item[-19:-13]}
            #reftDict = {r: item[-35:-20], "customer": item[-11:-5], "servicetag": item[-19:-13]}
            #stDict = {r: item[-35:-20], "customer": item[-11:-5], "servicetag": item[-19:-13]}
            #testDict.update({r:item[-19:-13]})
            #testDict.update({r:item[-11:-5]})
            r = r + 1
            #refNumList.append(item[-19:-13])
            #testDict ={}
           # serviceTagList.append(item[-11:-5])

        for k, v in custDict.items():
            print("cust dict",k, v)

        data = []
        data.append(customerInfoList)
        data.append(refNumList)
        data.append(serviceTagList)
        data.append(techName)

        print(data)
        data = pd.DataFrame(data).transpose()
        data.columns = ['Customer Name', 'Ref#', 'Service', 'Tech']

    def joinRootToCustomer(custName):
        print("in join fun",custName)
        rootPath = r'C:\Users\brian\Documents\Spreadsheets\Work_Instructions'
        searchPath = os.path.join(rootPath,custName)
        print(searchPath)
        #, os.path.join(sourceDir, sourceFileName))

app = QApplication(sys.argv)
mainwindow = EOS()
widget = QtWidgets.QStackedWidget()
widget.addWidget(mainwindow)
widget.setFixedHeight(850)
widget.setFixedWidth(1120)
widget.show()

try:
    sys.exit(app.exec_())
except:
    print("Exiting")