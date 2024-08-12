from openpyxl import load_workbook

from datetime import datetime, time
time_off_employees = []
all_employees = []

full_execuses = ['Annual Vacation EG','Work From Home EG','Work From Sahel','Special Occasion','Sick Leave']
partial_execuses = ['Half Day Work From Home EG','Half Day Vacation From Home EG']
lower_full_execuses = [execuse.lower() for execuse in full_execuses]
lower_partial_execuses = [execuse.lower() for execuse in partial_execuses]

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
        
        employee = {
            'name': row[0],
            'execuse': row[1],
            'id': row[2],
            'Sign In Time': None,
            'Sign Out Time': None,
        }
        time_off_employees.append(employee)
    return time_off_employees



# def validate_execuses():
    
#     employees = read_employees_with_timeoff()
#     for employee in employees:
#         if employee['execuse'].lower() not in lower_full_execuses and employee['execuse'].lower() not in lower_partial_execuses:
#             employee['execuse'] = 'invalid'
# validate_execuses() 

def exclude_employees_with_valid_execuses():
    wb = load_workbook('Attendance.xlsx')
    ws = wb.active
    to_be_excluded_employees = read_employees_with_timeoff()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if all(cell is None for cell in row):
            continue  # Skip empty rows
        if row[0] in [employee['id'] for employee in to_be_excluded_employees] :
            print('employee with id:',row[0],'is excluded')##################
exclude_employees_with_valid_execuses()

def read_all_employees():
    # Load the Excel file
    wb = load_workbook('Attendance.xlsx')
    ws = wb.active
    for row in ws.iter_rows(min_row=2, values_only=True):
        if all(cell is None for cell in row):
            continue  # Skip empty rows 
        employee = {
            'name': row[1],
            'id': row[0],
            'Sign In Time': None,
            'Sign Out Time': None,
        }
        all_employees.append(employee)
    return all_employees


def compare_sign_in_time_with_ten_oclock():
    # Load the Excel file
    wb = load_workbook('Attendance.xlsx')
    ws = wb['Working Sheet']  # Assuming this is the correct sheet name
    
    # Load all employees before comparing times
    read_all_employees()
    
    # Define 10:00:00 as a datetime.time object
    ten_oclock = time(10, 0, 0)
    
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        employee_id = row[0]  # Assuming employee ID is in the first column (A)
        excel_time_str = row[6]  # Assuming sign-in time is in the seventh column (G)

        # Check if the time is a string, and parse it
        if isinstance(excel_time_str, str):
            if excel_time_str == '#N/A':
                excel_time = None
            else:
                excel_time = datetime.strptime(excel_time_str, '%H:%M:%S').time()
        elif isinstance(excel_time_str, time):
            excel_time = excel_time_str
        else:
            excel_time = None  # Skip if the value is not a recognizable time format

        # Find the matching employee by ID in all_employees
        for employee in all_employees:
            if employee['id'] == employee_id:
                employee['Sign In Time'] = excel_time
                break  # Stop searching once the employee is found
        
        # Compare times
        if excel_time is not None:
            if excel_time < ten_oclock:
                print(f"Employee ID {employee_id}: Sign-in time {excel_time} is before 10:00:00")
            else:
                print(f"Employee ID {employee_id}: Sign-in time {excel_time} is after or at 10:00:00")

    print(all_employees)
compare_sign_in_time_with_ten_oclock()