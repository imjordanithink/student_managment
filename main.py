import datetime
from datetime import datetime as dt
import csv


################
# Musst learn either pickle or csv in order to continue, Feb 19, 2022
# Idea: end product would be a program that generates a .csv file with daily student list organized by belt and then subcatorigized by priority rating
# .csv file would then possibly be exported to google sheets or microsoft excel.
# Either learn how to make a weak api or stick to console operation. Cannot do both

student_list = []


#insert a preset futuredate into class __init__
time_change = datetime.timedelta(weeks= 8)
belt_map = {0:"No", 1:"White", 2:"Yellow", 3:"Orange", 4:"Green", 5:"Blue", 6:"Purple", 7:"Brown", 8:"Red", 9:"Black Stripe"}


# 1 --------- Student class
class Student:
    
    def __init__(self, first_name, last_name, belt, last_test_date, future_test_date, class_1, class_2 = "Mon 1:00AM", class_3 = "Mon 1:00AM"):
        self.first_name = first_name
        self.last_name = last_name
        self.username = last_name + first_name[0]
        self.belt = belt
        self.last_test_date = dt.strptime(last_test_date, "%d %b, %Y") # 1 Jan, 2022
        self.future_test_date = dt.strptime(future_test_date, "%d %b, %Y")
        self.class_1 = class_1 # Format: "%a %I:%M%p" == "Tue 6:15AM"
        self.class_2 = class_2 
        self.class_3 = class_3
        #A = 0-20 days till testing, B = 21-40 days till testing, C = 41-60 days till testing.
        self.priority = ""
    
    def check_testing(self):
        for student in student_list:
            priority_rating_raw =  self.future_test_date - dt.now() 
            priority_string = str(priority_rating_raw)
            pr = int(priority_string[0:3]) #pr = priority_rating
            if 40< pr <=60:
                self.priority = "C"
            elif 20< pr <=40:
                self.priority = "B"
            elif 0< pr <=20:
                self.priority = "A"

# 2 --------- Instructor class
class Inst:
    def __init__(self, first_name, last_name, username, password):
        self.first_name = first_name
        self.last_name = last_name
        self.username = username
        self.password = password

    def check_list(self, first_name, last_name):  
        for i in range(len(student_list)):
            for student in student_list:
                if first_name == student.first_name:
                    if last_name == student.last_name:
                        return False
                        break
                    else:
                        pass
                else:
                    pass
            return True
            
    def create_student(self):
        first_name = input("First name: ")
        last_name = input("Last name: ")
        if self.check_list(first_name, last_name) == True:
            belt = input("Belt rank: ")
            print("Ensure dates are entered in the following form: 1 Jan, 2022")
            last_test_date = input("Input Last Testing Date: ")
            future_test_date = input("Input Next Testing Date: ")
            class_1 = input("First class of the week (Tue 3:15PM): ")
            class_2 = input("Second class of the week (Sat 9:00AM). If none, press enter: ")
            class_3 = input("Third class off the week (Mon 6:00PM). If none, press eneter: ")
            student = Student(first_name, last_name, belt, last_test_date, future_test_date, class_1, class_2, class_3)
            student_list.append(student)
            file = open("student_dict.txt", "a")
            file.write("\n" + first_name + "|" + last_name + "|" + belt + "|" + last_test_date + "|" + future_test_date + "|" + class_1 + "|" + class_2 + "|" + class_3)
            file.close()
            print("{name} has sucessfully been added to the program".format(name = student.first_name))
        else:
            print("Student already exists.")

    def delete_student(self):
        first_name = input("First Name: ")
        last_name = input("Last Name: ")
        for student in student_list:
            if first_name == student.first_name:
                if last_name == student.last_name:
                    prompt = input("Are you sure you want to delete {name}'s data? (Y?N): ".format(name = student.first_name))
                    if prompt == "Y":
                        print("{name} has been removed.".format(name = student.first_name))
                        student_list.remove(student)
                    else:
                        print("Student does not exist.")
                        break
                else:
                    print("Student does not exist.")
                    pass
            else:
                print("Student does not exist.")
                pass
            

    
jordan = Student("Jordan", "Senko", "Red", "1 Jan, 2022", "1 Apr, 2022", "Tue 3:15pm")    

student_list = [jordan]
        



        
def main(inst):
    run = True
    while run == True:
        print(" ")
        print("--- MAIN PAGE --- {}".format(inst.username))
        print("""
        Ledger:
        CS = Create Student
        DS = Delete Student
        LO = Log Out
        """)
        x = input("Input Command: ")
        if x == "CS":
            run == False
            inst.create_student()
        elif x == "DS":
            run == False
            inst.delete_student()
        elif x == "CU":
            for instructor in instructor_list:
                print(instructor.username)
                print(instructor.password)
        elif x == "LO":
            quit()


def login():
    success = False
    #file = open("user_details.txt", "r")
    #for i in file:
        #a, b = i.split(",")
        #b = b.strip()
        #if (a==name and b == password):
            #success = True
            #break
    #file.close()
    username = input("Username: ")
    password = input("Password: ")
    index = 0
    for i in range(len(instructor_list)):
        index += 1
        for instructor in instructor_list:
            if instructor.username == username:
                if instructor.password == password:
                    success == True
                    break

    if success == True:
        print("Login Successful!")
        main(instructor_list[index-1])
    else:
        print("Incorrect login info")
        login()


def register():
    print(" ")
    username = input("Username: ")
    password = input("Password: ")
    first_name = input("First name: ")
    last_name = input("Last name: ")
    #file = open("user_details.txt", "a")
    #file.write("\n" + name + ", " + password)
    #file.close()
    instructor = Inst(first_name, last_name, username, password)
    instructor_list.append(instructor)
    welcoming_phrase = "{name} has been registered as an Instructor".format(name = first_name + "." + last_name[0])
    print(" ")
    print(welcoming_phrase)
    main(instructor)
instructor_list = []

def access(option):
    global name
    if (option == "login"):
        name = input("Enter your name: ")
        password = input("Enter your password: ")
        login(name, password)
    else:
        print("Enter your name and password to register")
        name = input("Enter your name: ")
        password = input("Enter your password: ")
        register(name, password)


def begin():
    global option
    print("--- Login/Register Page ---")
    option = input("Login or Register (log / reg): ")
    if option == "log":
        login()
    elif option == "reg":
        register()
    else:
        begin()

begin()
#access(option)








            

    


# wants to be able to see the students in classes that day sorted by their belt, and priority. 

# look into and learn CSV






#student.check_testing()
#print(student.priority)


