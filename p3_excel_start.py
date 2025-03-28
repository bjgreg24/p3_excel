# P3 EXCEL
# Brinley Gregory, Ethan Carn, Luke Miller
# Madi Diefenbach, Seth Mortenson, Sydney Trojahn Hedges
# IS 303 Section 004
# PROJECT DESCRIPTION: 

# from instructions
import openpyxl 
from openpyxl import Workbook
from openpyxl.styles import Font

# imports student class from studentClass.py for use in main program file
from studentClass import Student

print("1: Poorly_Organized_Data_1.xlsx")
print("2: Poorly_Organized_Data_2.xlsx")
testChoice = int(input("Choose which data you'd like to format (1 or 2): "))

if testChoice == 1:
    poorWorkbook = openpyxl.load_workbook("Poorly_Organized_Data_1.xlsx")
elif testChoice == 2:
    poorWorkbook = openpyxl.load_workbook("Poorly_Organized_Data_2.xlsx")

sheet = poorWorkbook.active

#########
# TO-DO #
#########

# Feel free to place your code here in the main file if you prefer that
# instead of making a function in a separate file.

# If you decide to define the function in a separate file, remember to import the function!
# You should also import the student class into your separate files!


# THIS FUNCTION SHOULD RETURN A LIST OF STUDENT OBJECTS FROM THE SELECTED EXCEL SHEET
studentList = getStudentObjects(sheet)

# THIS FUNCTION SHOULD CREATE A NEW EXCEL FILE, WITH SHEETS FOR EVERY CLASS (Algebra.xlsx, History.xlsx, etc.)
# SHOULD USE CLASS ATTRIBUTE IN STUDENT OBJECTS CONTAINED IN THE LIST
createWorksheets(studentList)

# THIS FUNCTION SHOULD ADD ALL THE STUDENT DATA TO THE NEW FILE AND CORRECT CLASS SHEETS
addStudentData(studentList)