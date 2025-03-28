# Luke Miller
# IS303 Section 004
# Creates a student class to be used in program files.

# Creates student class.
class Student:
    # Uses class name, student data string, and grade as parameters.
    # Key class attributes are first name, last name, student id, class name, and grade.
    def __init__(self, class_name, student_data, grade):
        self.class_name = class_name
        self.lname = ""
        self.fname = ""
        self.student_id = ""
        self.student_data = student_data
        self.unpackStudent(student_data)
        self.grade = grade

    # This method takes the string passed through student_data and splits it into each relevant attribute.
    # for instance, it will take "Miller" from "Miller_Luke_A512093" and add it to the last name attribute.
    def unpackStudent(self, student_data):
        dataList = student_data.split("_")
        self.fname = dataList[1]
        self.lname = dataList[0]
        self.student_id = dataList[2]


# THIS CODE IS ONLY TO TEST THE CLASS AND OBJECTS.
# RUN THIS FILE TO SEE HOW THE DATA IS STORED IN STUDENT OBJECTS.
# It will not be used in the final program!!
student_data = str(input("Enter student data string: "))
className = str(input("Enter class name: "))
grade = int(input("Enter grade: "))

oStudent = Student(className, student_data, grade)

print(oStudent.fname)
print(oStudent.lname)
print(oStudent.student_id)
print(oStudent.class_name)
print(oStudent.grade)