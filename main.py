from email import header
from itertools import count
from sqlite3 import Row
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
import datetime
from datetime import datetime as dt
import os
import re





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
def create_day_workbook():
    x = input("What day do you want to generate a sheet for? (Mon, Tue...): ")

    wbsd = load_workbook("Student_Data.xlsx")
    wssd = wbsd.active

    #wbop = Workbook()
    #wsop = wbop.active
    
    #wsop.title("{}.xlsx".format(str(dt.now())))

    temp_list = []
    count = 0
    for row in range(2, wssd.max_row + 1):
            for col in range(7, wssd.max_column + 1):
                char = get_column_letter(col)
                cell = char + str(row)
                print(cell)
                print(str(wssd[cell].value)[0:3])
                if wssd[cell].value == None:
                    wssd[cell] = "N/A"
                elif str(wssd[cell].value)[0:3] == x:
                    for col2 in range(1, 5):
                        char2 = get_column_letter(col2)
                
                        if count == 4:
                            temp_list.append(wssd[cell].value)
                            count = 0
                        else:
                            temp_list.append(wssd[char2 + str(row)].value)
                            count += 1
                
    
    key_list = []
    for i in range(4, len(temp_list), 5):
        key_list.append(temp_list[i])
        
    print(key_list)
        
     

    print(temp_list)
    temp_list = []

    wbsd.save("Student_Data.xlsx")

#for row in range(2, ws.max_row + 1):
        #first_name = ws["A" + str(row)].value
        #if first_name == f:
            #for row2 in range(2, ws.max_row + 1):
                    #last_name_cell = ws["B" + str(row2)].value
                    #if last_name_cell == l:
                        #display_student(row - 1)
                        #break
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
def update_priority_ratings():
    wb = load_workbook("Student_Data.xlsx")
    ws = wb.active

    for row in range(2, ws.max_row + 1):
        future_test_date = ws["F" + str(row)].value # "F" is Future Test Date column
        check_testing(future_test_date)
        ws["D" + str(row)] = priority # "D" is Priority Rating column

    wb.save("Student_Data.xlsx")

def check_testing(future_test_date):
    global priority
    
    priority_rating_raw = dt.strptime(future_test_date, "%d %b, %Y") - dt.now() 
    priority_string = str(priority_rating_raw)
    priority_string.replace("days,", "    ")
    pr_string = priority_string.split(" ")[0]
    pr = int(pr_string[0:3]) #pr = priority_rating
    if pr > 61:
        priority = "D"
    elif 40< pr <=60:
        priority = "C"
    elif 20< pr <=40:
        priority = "B"
    elif 0< pr <=20:
        priority = "A"

# --------------Regular Expressions Legend--------------
# CTRL + F = Search and replace bar
# MetaCharacters (Need to be escaped):
# .[{()\^$?*+
# \. <---- "." has been "escaped" so that we look for the literal of "." and not everything

# .       - Any Character Except New Line
# \d      - Digit (0-9)
# \D      - Not a Digit (0-9)
# \w      - Word Charecter (a-z, A-Z, 0-9, _)
# \W      - Not a Word Charecter
# \s      - Whitespace (space, tab, newline)
# \S      - Not a Whitespace

# --------- Anchors ---------------
# \b      - Word Boundry
# \B      - Not a Word Boundry
# ^       - Beginning of a String
# $       - End of a string

# []      - Matches characters IN brackets
# [^ ]    - Matches charecters NOT in brackets
# |       - Either or
# ( )     - Group

# --------- Quantifiers -----------
# *       - 0 or more
# +       - 1 or more
# ?       - 0 or 1
# {3}     - Exact number
# {3,4}   - Range of numbers (Min, Max)    
    
def create_student():
    first_name = input("Enter First Name: ")
    last_name = input("Enter Last Name: ")
    belt = input("Enter Belt Rank: ")
    last_test_date = input("Enter Last Test Date (1 Jan, 2022): ")
    if re.search(r"(\d{1,2}) (\w{3}), (\d{4})", last_test_date) == None:
        print("Incorrect Format. Try Again.")
        create_student()
    future_test_date = input("Enter Next Test Date (1 Apr, 2022): ")
    if re.search(r"(\d{1,2}) (\w{3}), (\d{4})", future_test_date) == None:
        print("Incorrect Format. Try Again.")
        create_student()
    first_class = input("Enter 1st Class (Mon 3:45PM): ")
    if re.search(r"(\w{3}) (\d{1,2}):(\d{0,2}") == None:
        print("Incorrect Format. Try Again.")
    second_class = input("Enter 2nd Class (Mon 3:45PM): ")
    if re.search(r"(\w{3}) (\d{1,2}):(\d{0,2}") == None:
        print("Incorrect Format. Try Again.")
    third_class = input("Enter 3rd Class (Mon 3:45PM): ")
    if re.search(r"(\w{3}) (\d{1,2}):(\d{0,2}") == None:
        print("Incorrect Format. Try Again.")
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

    print("{} {} has been added to the data sheet.".format(first_name, last_name))
    display()

def display_student(row_value):
    wb = load_workbook("Student_Data.xlsx")
    ws = wb.active
    count = 1
    header_dict = {1 : "First Name: ", 2 : "Last Name: ", 3 : "Belt Rank: ", 4 : "Priority Rating: ", 5 : "Last Test Date: ", 6 : "Next Test Date: ", 7 : "First Class: ", 8 : "Second Class: ", 9 : "Third Class: "}

    student_row_index = list(ws.rows)[row_value]
  
    for data in student_row_index:
        if data.value == None:
            print(header_dict[count] + "N/A")
            count += 1
        else:
            print(header_dict[count] + data.value)
            count += 1
    wb.save("Student_Data.xlsx")
    display()
          
def view_student():
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

def display():
    print("""
    ------ LEGEND ------
    vs = View Student Info
    vd = View the Day
    cs = Create Student Profile
    --------------------
    """)
    x = input("Input a Command: ")
    if x == "vs":
        view_student()
    elif x == "cs":
        create_student()
    elif x == "vd":
        create_day_workbook()

update_priority_ratings()       
display()

