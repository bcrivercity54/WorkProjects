import sys
import os
from PyQt5.uic import loadUi
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QDialog, QApplication
from pathlib import Path
import glob


class EOS(QDialog):
    def __init__(self):
        super(EOS, self).__init__()
        loadUi("test3.ui", self)
        self.tableWidget.setColumnWidth(0, 250)
        self.tableWidget.setColumnWidth(1, 250)
        self.tableWidget.setColumnWidth(2, 250)

        # print('test')

        self.grabButton.clicked.connect(self.getbwi)

        # self.loaddata()

    def getbwi(self=None):
        print("in the func")
        custtext = self.textCustomer.toPlainText()
        refText = self.refNum.toPlainText()

        print(custtext,refText)
        #joinRootToCustomer(custtext)

        rootPath = r'C:\Users\brian\Documents\Spreadsheets\Work_Instructions'

        searchPath = os.path.join(rootPath, custtext)
        print(searchPath)

        #

        # Search the folder for the entered ref number add to a list or dictionary

        # pull info from each WI and add to a dataframe
        fileList = ["GE_MR_OEM_4242234_9X3M2Y.xlsm", "GE_MR_OEM_4242234_8X3M2Y.xlsm", "GE_MR_OEM_4242234_7X3M2Y.xlsm"]
        # ../../Tools/*/*.
        #findFiles =
        #print("glob path",glob.glob(searchPath/+"4283348*"+".xlsm"))

        print("ending here")
        #[str(pp) for pp in glob_path.glob("**/*.txt")]
        #fileList =  [name for name in os.listdir(searchPath) if name.endswith(".xlsm")]
        #for i in fileList:

            #print(i)
        #print(glob_path("*4283348*.xlsm"))
        #for i in fileList:
            #print("I",i)

        #customerInfoList = []
        #refNumList = []
        #serviceTagList = []
        #techName = ['Brian', 'Brian', 'Brian']

        #for item in fileList:
        #    customerInfoList.append(item[-35:-20])
        #    refNumList.append(item[-19:-13])
        #    serviceTagList.append(item[-11:-5])

        #data = []
        #data.append(customerInfoList)
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
