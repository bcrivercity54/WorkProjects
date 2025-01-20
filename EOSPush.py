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
import pandas as pd

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
        #print("refText",type(refText))

        #print(custtext,refText)
        #joinRootToCustomer(custtext)

        rootPath = r'C:\Users\brian\Documents\Spreadsheets\Work_Instructions'

        searchPath = os.path.join(rootPath, custtext)
        #print("search path",searchPath)

        #pattern = "*refText*"
        #print("pattern",type(refText))
        # Search the folder for the entered ref number add to a list or dictionary

        fileList = []

        for root, dirs, files in os.walk(searchPath):
            #print("YABBA",pattern)
            if fnmatch.filter(files,"4283348"):
                fileList.append(files)


        fileList = ["GE_MR_OEM_4242234_9X3M2Y.xlsm", "GE_MR_OEM_4242234_8X3M2Y.xlsm", "GE_MR_OEM_4242234_7X3M2Y.xlsm"]

        mergelist = []
        customerInfoList = ["customer"]
        refNumList = []
        servicetaglist = []
        techName = ['Brian', 'Brian', 'Brian']
        keys = {"customer","refnum","servicetag"}
        wiDict = dict([(key,[]) for key in keys])
        print(wiDict)

        r = 0
        for item in fileList:
            #print("item",item)
            wiDict["customer"].append(item[-35:-20])
            wiDict["refnum"].append(item[-19:-13])
            wiDict["servicetag"].append(item[-11:-5])
            print(wiDict)
           # customerInfoList.append(item[-35:-20])
           #custdict.append()
            #refNumList.append(item[-19:-13])
            #servicetaglist.append(item[-11:-5])

        #mergelist = customerInfoList + refNumList + servicetaglist

        #for x in customerInfoList:
            #print("merglist",x)

        def loaddata(self,wiDict):
            print("in load data")
            row = 0
            self.tableWidget.setRowCount(len(wiDict))
            for wi in wiDict:
                self.tableWidget.setItem(row,0,QtWidgets.QTableWidgetItem(wi["customer"]))

            row = row+1

        #self.tableWidget.setRowCount(len(merglist))
        #for wi in mergelist:
            #self.tableWidget.setItem(row,0,QtWidgets.QTableWidgetItem(mergelist[""]))
        #for k, v in custDict.items():
            #print("cust dict",k, v)


        df = pd.DataFrame({'col':"FUCK"})
        print("test")
        print(df)
       # data = mergelist
        #data.append(refNumList)
       #data.append(serviceTagList)
        #data.append(techName)

        #print(data)
        #data = pd.DataFrame(data).transpose()
        #data.columns = ['Customer Name', 'Ref#', 'Service', 'Tech']

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