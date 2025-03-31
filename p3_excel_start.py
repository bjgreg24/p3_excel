# SEEING IF THIS WORKS

# P3 EXCEL
# Brinley Gregory, Ethan Carn, Luke Miller
# Madi Diefenbach, Seth Mortensen, Sydney Trojahn Hedges
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

# step 2 in to-do Use a for loop to get every row of data in the poorly formatted excel worksheet.
def getStudentObjects (sheet) :
    iRow = 2 
    studentList = []
    # create while loop to run as long as the row has data
    while sheet["A" + str(iRow)].value != None :
        # gather variables
        subject = sheet["A" + str(iRow)].value 
        studentData = sheet["B" + str(iRow)].value
        grade = sheet["C" + str(iRow)].value
        # create student object passing it the parameters of subject, studentData, and grade
        student = Student(subject, studentData, grade)
        # add student object to list of students
        studentList.append(student)
        # move to the next row
        iRow += 1
    # return the student list
    return studentList


# THIS FUNCTION SHOULD RETURN A LIST OF STUDENT OBJECTS FROM THE SELECTED EXCEL SHEET
studentList = getStudentObjects(sheet)

# THIS FUNCTION SHOULD CREATE A NEW EXCEL FILE, WITH SHEETS FOR EVERY CLASS (Algebra.xlsx, History.xlsx, etc.)
# SHOULD USE CLASS ATTRIBUTE IN STUDENT OBJECTS CONTAINED IN THE LIST

# create function that creates the worksheets and names them based on the subject
def createWorksheets (studentList) :
    # create an workbook to store organized data
    organizedWorkbook = Workbook () 
    # for each student in the list of students 
    for student in range(len(studentList)) :
        # if the subject name is not a worksheet name enter 
        if studentList[student].class_name not in organizedWorkbook.sheetnames :
            # creates a worksheet titles the subject
            organizedWorkbook.create_sheet(title=studentList[student].class_name)
            print(f"{studentList[student].class_name} was added as a worksheet.")
        # deletes the original worksheet
        if "Sheet" in organizedWorkbook.sheetnames :
            del organizedWorkbook["Sheet"]
    # saves the workbook 
    organizedWorkbook.save(filename="Organized_Data.xlsx")
    organizedWorkbook.close ()
    
createWorksheets(studentList)

# THIS FUNCTION SHOULD ADD ALL THE STUDENT DATA TO THE NEW FILE AND CORRECT CLASS SHEETS
#addStudentData(studentList)
addStudentData(studentList)

def addFilter (OrganizedWorkbook) :

    wb = openpyxl.load_workbook(OrganizedWorkbook)

    for worksheet in wb.worksheets :
        currWorksheet = wb[worksheet] 
        max_row = currWorksheet.max_row
        currWorksheet.auto_filter.ref = f"A1:D{max_row}"




# add filter function
def addFilter (OrganizedWorkbook) :
# load workbook
    wb = openpyxl.load_workbook(OrganizedWorkbook)
# for each worksheet in the workbook
    for worksheet in wb.worksheets :
        # make the current worksheet
        currWorksheet = wb[worksheet] 
        # find the max row
        max_row = currWorksheet.max_row
        # autofilter the worksheet
        currWorksheet.auto_filter.ref = f"A1:D{max_row}"




#Format colmns 
def format_columns(): 
    wb = openpyxl.load_workbook("Organized_Data.xlsx")
    cols = ['A', 'B', 'C', 'D', 'F', 'G']
    bold_font = Font(bold=True)
    for sheet in wb.worksheets:
        for col_letter in cols:
            cell = sheet[f'{col_letter}1']

            cell.font = bold_font

            header_text = str(cell.value) if cell.value else ""
            new_width = len(header_text) + 5
            sheet.column_dimensions[col_letter].width = new_width
            wb.save ("Organized_Data.xlsx")
