# P3 EXCEL
# ALL OUR NAMES:
# IS 303 
# Section 004
# PROJECT DESCRIPTION: 

# from instructions
import openpyxl 
from openpyxl import Workbook
from openpyxl.styles import Font

# to import existing workbook
import pandas as pd

df = pd.read_excel("Poorly_Organized_Data_1.xlsx")

# This is for the first example of poorly organized data
firstWorkbook = Workbook() #workbook object

currSheet = firstWorkbook.active #object for the object Workbook

firstWorkbook.remove(firstWorkbook["Sheet"])

classNames = [] # list for class names

for iCount in df.iloc[:,0]:
    if iCount not in classNames :
        classNames.append(iCount)

for items in classNames:
    firstWorkbook.create_sheet(classNames[items])

print(classNames)