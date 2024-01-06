import requests
from requests.auth import HTTPBasicAuth
from datetime import datetime

class Employee:
    def __init__(self, employee_id, name):
        self.employee_id = employee_id
        self.name = name

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
                
                print(f" Shift - In: {in_time[:-9]} {in_time[11:]}, Out - {out_time[:-9]} {out_time[11:]}, Total Time: {formatted_time}")
        else:
            print(f"Failed to retrieve shifts for employee {employee.name}. Status code:", response.status_code)

# Main execution
if __name__ == "__main__":
    # API key 
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
