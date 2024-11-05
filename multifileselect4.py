#from tkinter import *
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
#from tkinter.messagebox import showinfo
from openpyxl import Workbook
import openpyxl
from tkinter import messagebox 
#import sleep

### create the root window
##root = tk.Tk()
##root.title('Tkinter File Dialog')
##root.resizable(False, False)
##root.geometry('500x550')

rowInt = 0

outStringList = []
fileList = []

class App(tk.Tk):
    def __init__(self):
        super().__init__()

        print("in init")
        self.geometry("500x580")

        self.row_var = tk.StringVar()

        self.col_var = tk.StringVar()

        self.columnconfigure(0,weight = 1)
        self.columnconfigure(1,weight = 1)
        self.columnconfigure(2,weight = 1)

        

        self.outStringList = outStringList

        self.put_Widgets()
        #self.updateOutLabel()
        

    def put_Widgets(self):

        padding = {'padx': 5, 'pady': 5}
        ####   txt= file.read()
        ttk.Label(self,text="Row Location:").grid(column = 0 ,row=0,**padding)
        ttk.Label(self,text="Column Offset:").grid(column = 0 ,row=1,**padding)
        
        
        rowText = ttk.Entry(self,textvariable = self.row_var)
        rowText.grid(column=1,row=0, **padding)
        rowText.focus()

        colText = ttk.Entry(self,textvariable = self.col_var)
        colText.grid(column=1,row=1, **padding)
        #colText.focus()
        

        open_button = ttk.Button(self,text='Select Files',command=select_files)
        open_button.grid(column=1,row=5, **padding)

        resultsButton = ttk.Button(self,text="Grab results")
        resultsButton.grid(column = 1, row=15)

        updateOutLabel = ttk.Label(text="Results")
        updateOutLabel.grid(column = 1, row = 30)

        rowString = int(self.row_var.get())
        rowInt = int(rowString.strip())

        messagebox.showinfo("showinfo", "Information")


    def outLabel(self):
        print("test")

def select_files():
    filetypes = (('Work Instructions files', '*.xlsx'),('All files', '*.*'))   

    path = fd.askopenfilenames(title='Open files',initialdir='C:\\Users\\brian\\Documents\\Spreadsheets\\Work_Instructions\\GE_HC_OEM',filetypes=filetypes)
    fileList.append(path)
    for i in fileList:
        file = i
        findMissedSignOffs(file)
        
   
def findMissedSignOffs(fileList):

    #outStringList = []
     
    for x in fileList:
               
        wb = openpyxl.load_workbook(x, read_only=False)
        ws = wb.active
        for row in ws.iter_rows(rowInt):
             for cell in row:
                 #print("cell value",cell.value)
                 if cell.value == "Tech Initial:" and cell.offset(row=0, column=1).value is  None:
                     #print("cell value",cell.value)
                     #rowNum ="row("+row+")"
                     rowNum = str((row[0:1]))
                     sRowNum = rowNum.replace("<Cell 'Sheet1'.",'')
                     sRowNum = sRowNum.replace(">,",'')
                     #print("Row num",sRowNum)
                     #print(row,cell.value,cell.offset(row=0, column=1).value)
                     outStringList.append(sRowNum) #label.config(text=sRowNum, font=('Courier 13 bold'))
     
    
   
    
        
results = App()
#print("duh",t)


      

if __name__ == "__main__":
    app = App()
    app.mainloop()
