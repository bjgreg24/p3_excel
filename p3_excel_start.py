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
def addStudentData(studentList, organizedWorkbook):
    # open workbook
    wb = openpyxl.load_workbook(organizedWorkbook)

    # create headers
    headers = ["Last Name", "First Name", "Student ID", "Grade"]

    # loop through every sheet in the workbook
    for sheet_name in wb.sheetnames:
        # activate worksheet
        worksheet = wb[sheet_name]
        wb.active = worksheet  

        # put headers on every sheet
        for col, header in enumerate(headers, start=1):
            worksheet.cell(row=1, column=col, value=header)
        
        # loop through every student object and see if they are in that class
        # if student is in the class, input their data into the sheet
        for student in range(len(studentList)):
            if wb.active.title == studentList[student].class_name:
                iRow = wb.active.max_row + 1  # Get the first empty row
                iCol = "A"
                wb.active[iCol + str(iRow)].value = studentList[student].lname
                iCol = "B"
                wb.active[iCol + str(iRow)].value = studentList[student].fname
                iCol = "C"
                wb.active[iCol + str(iRow)].value =studentList[student].student_id
                iCol = "D"
                wb.active[iCol + str(iRow)].value = studentList[student].grade

    # saves the workbook 
    wb.save(organizedWorkbook)
    wb.close()


organizedWorkbook = "Organized_Data.xlsx"
addStudentData(studentList,organizedWorkbook)