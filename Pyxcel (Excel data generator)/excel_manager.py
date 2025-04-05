import openpyxl
from openpyxl import load_workbook
import data_creater


def generate_data(num_rows, file_path):
    # Create a new workbook and worksheet
    wb = load_workbook(file_path)  # Load the existing workbook
    ws = wb.active # Get the active worksheet
    ws.title = "Data" # Set the title of the worksheet

    # Add headers to the worksheet
    ws.append(["Serial No", "Name", "Address", "Email", "Phone Number", "Date of Birth", "Job", "Company"])

    # Generate  data with serial numbers
    for i in range(1, num_rows + 1):  # Start serial numbers from 1
        ws.append([
            i,  # Serial number
            data_creater.name,  # Name
            data_creater.address,  # Address
            data_creater.email,  # Email
            data_creater.phone_number,  # Phone number
            data_creater.date_of_birth,  # Date of birth
            data_creater.job,  # Job
            data_creater.company  # Company
        ])

    # Save the workbook to the specified file path
    try:
        wb.save(file_path)  # Save the workbook
        return True  # Indicate success
    except Exception as e: # Handle any exceptions that occur during saving
        print(f"Error saving file: {e}")
        return False  # Indicate failure
    