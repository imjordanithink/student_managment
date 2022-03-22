from email import header
from itertools import count
from sqlite3 import Row
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
import datetime
from datetime import datetime as dt





data = {
    "Jordan":{
        "Last Name": "Senko",
        "Belt": "Red",
        "Last Test Date": "1 Jan, 2022",
        "Next Test Date": "1 Apr, 2022",
        "First Class": "Mon 6:15PM",
        "Second Class": "Tue 7:45PM",
        "Third Class": "Wed 5:45PM"
    },
    "Nikki":{
         "Last Name": "Bush",
        "Belt": "Red",
        "Last Test Date": "1 Jan, 2022",
        "Next Test Date": "1 Apr, 2022",
        "First Class": "Mon 6:15PM",
        "Second Class": "Tue 7:45PM",
        "Third Class": "Wed 5:45PM"
    },
    "Mark":{
         "Last Name": "Bush",
        "Belt": "Green",
        "Last Test Date": "1 Jan, 2022",
        "Next Test Date": "1 Apr, 2022",
        "First Class": "Mon 6:15PM",
        "Second Class": "Tue 7:45PM",
        "Third Class": "Wed 5:45PM"
    }
}

#wb = Workbook()
#ws = wb.active
#ws.title = "Data"

#headings = ["Name"] + list(data["Jordan"].keys())
#ws.append(headings)

#for person in data:
    #points = list(data[person].values())
    #ws.append([person] + points)



#wb.save("Student_Data.xlsx")

#print(ws.max_row)
#print(ws.max_column)



def check_testing(future_test_date):
    global priority
    priority_rating_raw = dt.strptime(future_test_date, "%d %b, %Y") - dt.now() 
    priority_string = str(priority_rating_raw)
    pr = int(priority_string[0:3]) #pr = priority_rating
    if 40< pr <=60:
        priority = "C"
    elif 20< pr <=40:
        priority = "B"
    elif 0< pr <=20:
        priority = "A"
    


def create_student():
    first_name = input("Enter First Name: ")
    last_name = input("Enter Last Name: ")
    belt = input("Enter Belt Rank: ")
    last_test_date = input("Enter Last Test Date (1 Jan, 2022): ")
    future_test_date = input("Enter Next Test Date (1 Apr, 2022): ")
    first_class = input("Enter 1st Class (Mon 3:45PM): ")
    second_class = input("Enter 2nd Class (Mon 3:45PM): ")
    third_class = input("Enter 3rd Class (Mon 3:45PM): ")
    check_testing(future_test_date)

    temp_dict = {
        first_name:{
            "Last Name" : last_name,
            "Belt" : belt,
            "Priority" : priority,
            "Last Test Date" : last_test_date,
            "Future Test Date" : future_test_date,
            "First Class" : first_class,
            "Second Class" : second_class,
            "Third Class" : third_class
            }
    }

    wb = load_workbook("Student_Data.xlsx")
    ws = wb.active

    for person in temp_dict:
        points = list(temp_dict[person].values())
        ws.append([person] + points)
    
    wb.save("Student_Data.xlsx")





def display_student(row_value):
    wb = load_workbook("Student_Data.xlsx")
    ws = wb.active
    count = 1
    header_dict = {1 : "First Name: ", 2 : "Last Name: ", 3 : "Belt Rank: ", 4 : "Last Test Date: ", 5 : "Next Test Date: ", 6 : "First Class: ", 7 : "Second Class: ", 8 : "Third Class: "}

    student_row_index = list(ws.rows)[row_value]
    for data in student_row_index:
        print(header_dict[count] + data.value)
        count += 1
    wb.save("Student_Data.xlsx")
    

def view_student():
    found = False
    wb = load_workbook("Student_Data.xlsx")
    ws = wb.active

    x = input("Input Name (Jordan Senko): ")
    f, l = x.split(" ")
    
    for row in range(2, ws.max_row + 1):
        first_name = ws["A" + str(row)].value
        if first_name == f:
            for row2 in range(2, ws.max_row + 1):
                    last_name_cell = ws["B" + str(row2)].value
                    if last_name_cell == l:
                        display_student(row - 1)
                        break
    wb.save("Student_Data.xlsx")
            
            

#view_student()
create_student()
