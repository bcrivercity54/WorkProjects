import tkinter as tk
from tkinter import ttk
from tkinter.messagebox import showerror
from tkinter import filedialog as fd
from openpyxl import Workbook
import openpyxl
from openpyxl.cell import MergedCell

outStringList = []



def findMissedSignOffs(st,xFile):
        print("in missed signoffs")
        

        #row = int(self.rowNum.get())
        #print(type(row))
        #outStringList = []  
     
        #for x in fileList:
        wb = openpyxl.load_workbook(xFile, read_only=False)
        sheet_ranges = wb['Sheet1']
        ws = wb.active
        
        for row in ws.iter_rows(2):
            for cell in row:
                #print("cell value",cell.value)
                if bool(sheet_ranges.merged_cells.ranges):
                    if cell.value == "Tech Initial:" and cell.offset(row=0, column=1).value is  None:
                        #print("cell value",cell.value)
                        #rowNum ="row("+row+")"
                        rowNum = str((row[0:1]))
                        sRowNum = rowNum.replace("<Cell 'Sheet1'.",'')
                        sRowNum = sRowNum.replace(">,",'')
                        sRowNum = st + " " + sRowNum
                        print(sRowNum)
                        outStringList.append(sRowNum) #label.config(text=sRowNum, font=('Courier 13 bold'))
                        #showResults(outStringList)
##                        getResultsList = WIFrame.returnResults
##                        getResultsList(outStringList)
                        



     
            
class WIFrame(ttk.Frame):
    def __init__(self, container):
        super().__init__(container)


        self.outStringList = outStringList
        
        # field options
        options = {'padx': 5, 'pady': 5}

        # row label
        self.rowlabel = ttk.Label(self, text='Row')
        self.rowlabel.grid(column=0, row=0, sticky=tk.W, **options)

        self.collabel = ttk.Label(self, text='Col')
        self.collabel.grid(column=0, row=1, sticky=tk.W, **options)

        # entry
        self.rowNum = tk.StringVar()
        self.rowNum = ttk.Entry(self,width=3,textvariable=self.rowNum)
        self.rowNum.grid(column=1, row=0, **options)
        self.rowNum.focus()

        self.colNum = tk.StringVar()
        self.colNum = ttk.Entry(self,width=3,textvariable=self.colNum)
        self.colNum.grid(column=1, row=1, **options)

        self.select_button = ttk.Button(self, text='Select Files')
        self.select_button['command'] = self.select_files
        self.select_button.grid(column=2, row=5, sticky=tk.W, **options)

        self.grab_button = ttk.Button(self, text='Grab em')
        self.grab_button['command'] = self.returnResults
        self.grab_button.grid(column=2, row=15, sticky=tk.W, **options)

        # result label
        self.result_label = ttk.Label(self)
        self.result_label.grid(row=25,column=2,columnspan=3, **options)

        # add padding to the frame and show it
        self.grid(padx=10, pady=10, sticky=tk.NSEW)

    def select_files(self):

        fileList = []
     
        fileTuple = ()
        try:
            filetypes = (('Work Instructions files', '*.xlsx'),('All files', '*.*'))

            path = fd.askopenfilenames(title='Open files',initialdir='C:\\Users\\brian\\Documents\\Spreadsheets\\Work_Instructions\\GE_HC_OEM',filetypes=filetypes)
            
            #print(type(fileList))
            fileList = list(path)
         
            for file in fileList:
                serialTag = file[-11:-5]
                print(serialTag)                
                findMissedSignOffs(serialTag,file)
                
        except ValueError as error:
            showerror(title='Error', message=error)

    def returnResults(self):

        self.result_label.config(text=("\n".join(outStringList)))
    

class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title('Missed WI signoffs')
        self.geometry('300x370')
        self.resizable(False, False)


if __name__ == "__main__":
    app = App()
    WIFrame(app)
    app.mainloop()
