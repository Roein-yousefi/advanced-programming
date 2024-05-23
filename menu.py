import openpyxl
from input_user import InputUser

class Menu(InputUser):
    def __init__(self, workbook):
        self.workbook = workbook



    def show_menu(self):
        print("""
    1-add student
    2-delete student
    3-create sheets
    4-remove sheets
    5-best of student
    6-edit_student
    """)
    
    def choice(self):
        while True:
            try:
                choice = int(input('enter your choice: '))
                break
            except ValueError:
                print('Invalid input. Please enter a number.')
        
        if choice == 1 :
            InputUser.add_student(self)
        if choice == 2 :
            InputUser.delete_student(self)
        if choice == 3 :
            InputUser.create_sheets(self)
        if choice == 4 :
            InputUser.remove_sheets(self)
        if choice == 5 :
            InputUser.best_of_student(self)
        if choice == 6 :
            InputUser.edit_student(self)
        




workbook = openpyxl.load_workbook('students.xlsx')
clas = Menu(workbook)
clas.show_menu()
clas.choice()