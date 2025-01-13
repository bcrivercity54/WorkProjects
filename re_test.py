import os
import re
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd

string = r'C:\Users\brian\Documents\Spreadsheets\Work_Instructions\GE_HC_OEM'

refText = 4283348
print("reft",refText)
#srchString = refText

pair = re.compile("\d\d\d\d\d\d\d")


x = pair.search(string,refText)
print(type(x))
##print(x,refText)
print(x.group())

#'print(pair.search(srchString)
