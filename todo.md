PROGRAM FLOW

INPUT & PROCESSING

1: Set up initial program requirements.
- This should prepare the poorly formatted excel file to be used in the for loop shown below.

2: Use a for loop to get every row of data in the poorly formatted excel worksheet.
- Should get variables "class_name", "student_data", and "grade" (one variable for each column in the row.)

3: Create a student object based on each row of data.
- Define a student class. Student objects should have "first_name", "last_name", "student_id", "class_name", and "grade" as attributes.
- Student class should have a method to split the "student_data" string into the "first_name", "last_name", and "student_id" attributes when the object is created.
- Creates the student object in the loop using variables "class_name", "student_data", and "grade" from step 2 as object parameters.

4: Add each student object to list in the loop to pull from later when creating the new worksheet.
- At the end of the loop, you should have a list of unique student objects based on the data in the original worksheet.

OUTPUT

5: Creates a new output worksheet for each college class.
- On the first row of the new worksheet, create data labels for each data category.

6: Using the list of student objects, add data to new worksheet.
- Use student object attributes (such as first and last name) and insert the data into each relevant column in the new worksheet.

(Step 5 and 6 could be a separate file called output.py which defines a function that takes a list of student objects as a parameter that can be called in the main program file).