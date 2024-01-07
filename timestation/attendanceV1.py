import requests
from requests.auth import HTTPBasicAuth
from datetime import datetime, timedelta
import openpyxl
import os

class Employee:
    def __init__(self, employee_id, name):
        self.employee_id = employee_id
        self.name = name
        self.shifts ={}
        
    def add_shift(self, date, in_time, out_time, total_time):
        self.shifts[date] = [in_time, out_time, total_time]
        
def list_employees(api_key):
    # API endpoint for listing employees
    url = "https://api.mytimestation.com/v1.2/employees"

    # Make the GET request with HTTP Basic Authentication
    response = requests.get(url, auth=HTTPBasicAuth(api_key, ''))

    # Create and return a list of Employee objects if the request is successful
    if response.status_code == 200:
        employees_data = response.json()
        return [Employee(employee['employee_id'], employee['name']) for employee in employees_data['employees']]
    else:
        print("Failed to retrieve employees. Status code:", response.status_code)
        return []

def get_hours_worked_report(api_key, employees, start_date, end_date):
    # API endpoint for fetching shifts
    base_url = "https://api.mytimestation.com/v1.2/shifts"
    
    def minutes_to_hhmm(total_minutes):
        hours = total_minutes // 60
        minutes = total_minutes % 60
        return f"{hours:02d}:{minutes:02d}"
    
    def convert_datetime_format(date_string):
    # Parse the string into a datetime object
        dt = datetime.strptime(date_string, "%Y-%m-%d %H:%M:%S")
        
        # Format the datetime object into the desired format
        formatted_date = dt.strftime("%m/%d/%Y %H:%M")
        
        return formatted_date

    # Loop through each Employee object and fetch their shifts
    for employee in employees:
        # Construct the full URL with parameters
        full_url = f"{base_url}?employee_id={employee.employee_id}&start_date={start_date}&end_date={end_date}"

        # Make the GET request with HTTP Basic Authentication
        response = requests.get(full_url, auth=HTTPBasicAuth(api_key, ''))

        # Check if the request was successful
        if response.status_code == 200:
            # Print the employee name and their shift data
            print(f"\nName: {employee.name}")
             # Extract and print total_minutes and in/out status for each shift
            shifts = response.json().get('shifts', [])
            for shift in shifts:
                # Formating the dict data
                total_minutes = shift.get('total_minutes')
                formatted_time = minutes_to_hhmm(int(total_minutes))
                
                in_status = shift.get('in')
                in_time = next(iter(in_status.items()))[1] if in_status else 'N/A'
                if in_time != 'N/A':
                    in_time = in_time[:-6]
                    
                out_status = shift.get('out')
                out_time = next(iter(out_status.items()))[1] if in_status else 'N/A'
                if in_time != 'N/A':
                    out_time = out_time[:-6]
                    
                # there's a better way to do this but this works
                date =convert_datetime_format(in_time[:-9]+" "+in_time[11:])[:-6]
                clockin = convert_datetime_format(in_time[:-9]+" "+in_time[11:])[11:]
                clockout = convert_datetime_format(out_time[:-9]+" "+out_time[11:])[11:] 
                
                employee.add_shift(date,clockin,clockout,formatted_time)
                
            for shift in employee.shifts:
                print(shift, employee.shifts[shift])
        else:
            print(f"Failed to retrieve shifts for employee {employee.name}. Status code:", response.status_code)


def fill_missing_dates(employee,start_date):
    
    def convert_datetime_format(date_string):
    # Parse the string into a datetime object
        dt = datetime.strptime(date_string, "%Y-%m-%d")
        # Format the datetime object into the desired format
        formatted_date = dt.strftime("%m/%d/%Y")
        return formatted_date
    
    # Convert start and end dates to datetime objects
    start = datetime.strptime(convert_datetime_format(start_date), "%m/%d/%Y")
    end = datetime.strptime(max(employee.shifts.keys()), "%m/%d/%Y")

    # Initialize a date variable with the start date
    current_date = start

    # Iterate over the range of dates
    while current_date <= end:
        # Format the current date as MM/DD/YYYY
        date_str = current_date.strftime("%m/%d/%Y")

        # If the date is not in the dictionary, add it with a default value
        if date_str not in employee.shifts:
            employee.shifts[date_str] = [0, 0, 0]

        # Move to the next day
        current_date += timedelta(days=1)
    def sort_shift_data(shift_data):
        # Sort the dictionary by date keys in ascending order
        sorted_data = dict(sorted(shift_data.items(), key=lambda item: datetime.strptime(item[0], "%m/%d/%Y")))
        return sorted_data
    employee.shifts=sort_shift_data(employee.shifts)


def find_and_open_excel_files(employee ,folder_path):
    
    def get_month_name(date_str):
        # Parse the date string into a datetime object
        date_obj = datetime.strptime(date_str, '%m/%d/%Y')

        # Extract the month name and year
        month_name = date_obj.strftime('%b')
        year = date_obj.strftime('%y')

        # Return the formatted string
        return f"{month_name}.{year}"
    # List all .xlsx files in the given folder
    excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]

    search_string = employee.name
    
    sheet_name = get_month_name(next(iter(employee.shifts)))

    # Check each file to see if it contains the search string in its name
    fileFound=False#Dumb perfectionism
    for file in excel_files:
        if search_string in file:
            # Construct the full file path
            file_path = os.path.join(folder_path, file)

            # Open the workbook - add your logic here as needed
            workbook = openpyxl.load_workbook(file_path)
            if sheet_name not in workbook.sheetnames:
                raise ValueError(f"Sheet '{sheet_name}' not found in the workbook.")
            sheet = workbook[sheet_name]

            # Start from row 7, column 3 (which is 'C7')
            row = 7
            date_column = 1
            in_column = 2
            out_column = 3

            # Iterate over the data list and write to the cells
            for shift in employee.shifts:
                date_cell = sheet.cell(row=row, column=date_column)
                date_cell.value = shift
                in_cell = sheet.cell(row=row, column=in_column)
                in_cell.value = employee.shifts[shift][0]
                out_cell = sheet.cell(row=row, column=out_column)
                out_cell.value = employee.shifts[shift][1]
                row += 1  # Move to the next row

            # Save the workbook
            workbook.save(file_path)

            # Example: Print the sheet names
            print(f"Updated file for {employee.name}")
            fileFound=True#Dumb perfectionism
            break
    if fileFound != True:
        print(f"No file found for {employee.name}")

# Main execution
if __name__ == "__main__":
    
    # API key should be saved in a separate .env file stored locally
    with open("timestation\\APIKEY.env", 'r') as file:
        api_key = file.read()

    # Calculate the first day of the current month
    today = datetime.now()
    start_date = today.replace(day=1).strftime("%Y-%m-%d")

    # End date is today's date
    end_date = today.strftime("%Y-%m-%d")

    # Get the list of Employee objects
    employees = list_employees(api_key)

    # Get the hours worked report for each employee
    if employees:
        get_hours_worked_report(api_key, employees, start_date, end_date)
        for employee in employees:
            if employee.shifts:
                fill_missing_dates(employee, start_date)
                find_and_open_excel_files(employee, "timestation")
        print("Success!")
