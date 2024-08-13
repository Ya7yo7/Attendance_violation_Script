from openpyxl import load_workbook

from datetime import datetime, time
time_off_employees = []
all_employees = []
employee_ids = []
employees_dict = {}
accepted_execuses = ['Annual Vacation EG', 'Work From Home EG', 'Work From Sahel', 'Special occasion', 'Sick leave']
def read_ids():
    # Load the Excel file
    wb = load_workbook('Attendance.xlsx')
    ws = wb.active
    for row in ws.iter_rows(min_row=2, values_only=True):
        if all(cell is None for cell in row):
            continue  # Skip empty rows 
        employee_ids.append(row[0])
class Employee:
    def __init__(self, name, execuse, id, sign_in_time, sign_out_time,sign_in_violation ,sign_out_violation , duration_violation):
        self.name = name
        self.execuse = execuse
        self.id = id
        self.sign_in_time = sign_in_time
        self.sign_out_time = sign_out_time
        self.sign_in_violation = sign_in_violation
        self.sign_out_violation = sign_out_violation
        self.duration_violation = duration_violation
        
    def __repr__(self):
        return f"Employee({self.name}, {self.execuse}, {self.id}, {self.sign_in_time}, {self.sign_out_time}, {self.sign_in_violation}, {self.sign_out_violation}, {self.duration_violation})"

def read_all_employees():
    # Load the Excel file
    wb = load_workbook('Attendance.xlsx')
    ws = wb.active
    for row in ws.iter_rows(min_row=2, values_only=True):
        if all(cell is None for cell in row):
            continue  # Skip empty rows 
        employee = Employee(row[1],None,row[0],row[6],row[9],False,False,False)
        all_employees.append(employee)
        employees_dict[employee.id] = employee

def read_employees_with_timeoff():
    # Load the Excel file
    wb = load_workbook('TimeOffDay.xlsx')
    sheet = wb['Sheet1']
    # Get the worksheet
    ws = wb.active
    # Read the data from the worksheet
    for row in ws.iter_rows(min_row=2, values_only=True):
        if all(cell is None for cell in row):
            continue  # Skip empty rows
        employees_dict[row[2]].execuse = row[1]

def check_sign_in_violations():
    ten_oclock = time(10, 5, 0)
    one_oclock = time(13, 5, 0)
    for employee in employees_dict.values():
        if isinstance(employee.sign_in_time, str):
            if employee.sign_in_time == '#N/A':
                employee.sign_in_time = None
                if employee.execuse in accepted_execuses:
                    employee.sign_in_violation = False
                elif employee.execuse == None:
                    employee.sign_in_violation = True
                continue
            else:
                employee.sign_in_time = datetime.strptime(employee.sign_in_time, '%H:%M:%S')
        
        if employee.execuse in accepted_execuses:
            employee.sign_in_violation = False
            continue
            
        if employee.execuse == 'Half Day Work From Home EG':
            if employee.sign_in_time > one_oclock:
                employee.sign_in_violation = True
        elif employee.execuse == 'Half Day Vacation From Home EG':
            if employee.sign_in_time > ten_oclock:
                employee.sign_in_violation = True
        elif employee.execuse == None:
            if employee.sign_in_time > ten_oclock:
                employee.sign_in_violation = True
def check_sign_out_violations():
    four_oclock = time(15, 55, 0)
    one_oclock = time(12, 55, 0)

    for employee in employees_dict.values():
        if isinstance(employee.sign_out_time, str):
            if employee.sign_out_time == '#N/A':
                employee.sign_out_time = None
                if employee.execuse in accepted_execuses:
                    employee.sign_out_violation = False
                elif employee.execuse == None:
                    employee.sign_out_violation = True
                continue
            else:
                employee.sign_out_time = datetime.strptime(employee.sign_out_time, '%H:%M:%S')

        if employee.execuse == 'Annual Vacation EG' or employee.execuse == 'Work From Home EG' or employee.execuse == 'Work From Sahel'\
        or employee.execuse == 'Special occasion' or employee.execuse == 'Sick Leave':
             employee.sign_out_violation = True
             continue
       
        if employee.execuse == 'Half Day Work From Home EG':
            if employee.sign_out_time < four_oclock:
                employee.sign_out_violation = True
        elif employee.execuse == 'Half Day Vacation From Home EG':
            if employee.sign_out_time < one_oclock:
                employee.sign_out_violation = True
        elif employee.execuse == None:
            if employee.sign_out_time < four_oclock:
                employee.sign_out_violation = True
                
def NoShowNoExcuse():
    for employee in employees_dict.values():
        # print(employee.sign_in_time)
        if employee.sign_in_time == None and employee.sign_out_time == None and not employee.execuse:
            print(f"Employee {employee.name} with ID {employee.id} is absent with no excuse.")        



def main():
    read_ids()
    read_all_employees()
    read_employees_with_timeoff()
    check_sign_in_violations()
    check_sign_out_violations()
    for employee in employees_dict.values():
        print(employee,'\n')

    NoShowNoExcuse()
    # print(all_employees)
    # read_employees_with_timeoff()
    # print(time_off_employees)
    # exclude_employees_with_valid_execuses()
    # compare_sign_in_time_with_ten_oclock()
    # print(all_employees)
    # print(time_off_employees)
    # print(employees_dict)

if __name__ == '__main__':
    main()