from openpyxl import load_workbook

from datetime import datetime, time
time_off_employees = []
all_employees = []
employee_ids = [10022, 10025, 10035, 10037, 10046, 10052, 10073, 10119, 10123, 10249, 10100]
employees_dict = {}
full_execuses = ['Annual Vacation EG','Work From Home EG','Work From Sahel','Special Occasion','Sick Leave']
partial_execuses = ['Half Day Work From Home EG','Half Day Vacation From Home EG']
lower_full_execuses = [execuse.lower() for execuse in full_execuses]
lower_partial_execuses = [execuse.lower() for execuse in partial_execuses]

class Employee:
    def __init__(self, name, execuse, id, sign_in_time, sign_out_time,sing_in_violation = False,sign_out_violation = False):
        self.name = name
        self.execuse = execuse
        self.id = id
        self.sign_in_time = sign_in_time
        self.sign_out_time = sign_out_time
        self.sing_in_violation = sing_in_violation
        self.sign_out_violation = sign_out_violation
        
    def __str__(self):
        return f"Employee {self.name} with ID {self.id} has execuse {self.execuse} and signed in at {self.sign_in_time} and signed out at {self.sign_out_time}"

    def __repr__(self):
        return f"Employee({self.name}, {self.execuse}, {self.id}, {self.sign_in_time}, {self.sign_out_time})"

def read_all_employees():
    # Load the Excel file
    wb = load_workbook('Attendance.xlsx')
    ws = wb.active
    for row in ws.iter_rows(min_row=2, values_only=True):
        if all(cell is None for cell in row):
            continue  # Skip empty rows 
        employee = Employee(row[1],None,row[0],row[6],row[9])
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



# def compare_sign_in_time_with_ten_oclock():
#     # Load the Excel file
#     wb = load_workbook('Attendance.xlsx')
#     ws = wb['Working Sheet']  # Assuming this is the correct sheet name
    
#     # Load all employees before comparing times
#     read_all_employees()
    
#     # Define 10:00:00 as a datetime.time object
#     ten_oclock = time(10, 0, 0)
    
#     for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
#         employee_id = row[0]  # Assuming employee ID is in the first column (A)
#         excel_time_str = row[6]  # Assuming sign-in time is in the seventh column (G)

#         # Check if the time is a string, and parse it
#         if isinstance(excel_time_str, str):
#             if excel_time_str == '#N/A':
#                 excel_time = None
#             else:
#                 excel_time = datetime.strptime(excel_time_str, '%H:%M:%S').time()
#         elif isinstance(excel_time_str, time):
#             excel_time = excel_time_str
#         else:
#             excel_time = None  # Skip if the value is not a recognizable time format

#         # Find the matching employee by ID in all_employees
#         for employee in all_employees:
#             if employee['id'] == employee_id:
#                 employee['Sign In Time'] = excel_time
#                 break  # Stop searching once the employee is found
        
#         # Compare times
#         if excel_time is not None:
#             if excel_time < ten_oclock:
#                 print(f"Employee ID {employee_id}: Sign-in time {excel_time} is before 10:00:00")
#             else:
#                 print(f"Employee ID {employee_id}: Sign-in time {excel_time} is after or at 10:00:00")

#     print(all_employees)
# compare_sign_in_time_with_ten_oclock()



def main():
    read_all_employees()
    print(employees_dict)
    print('---------------------------------')
    read_employees_with_timeoff()
    print(employees_dict)
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