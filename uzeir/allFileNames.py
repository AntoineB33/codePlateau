import os
from openpyxl import Workbook

def list_files_in_folder(folder_path):
    # List all files in the folder
    files = os.listdir(folder_path)
    
    # Remove extensions from file names
    file_names = [os.path.splitext(file)[0] for file in files]
    
    return file_names

def write_file_names_to_excel(file_names, excel_path):
    # Create a new Excel workbook and select the active worksheet
    wb = Workbook()
    ws = wb.active
    
    # Write file names to the first column
    for idx, name in enumerate(file_names, start=1):
        ws.cell(row=idx, column=1, value=name)
    
    # Save the workbook
    wb.save(excel_path)

# Example usage
folder_path = r"C:\path\to\your\folder"
excel_path = r"C:\path\to\your\output.xlsx"

file_names = list_files_in_folder(folder_path)
write_file_names_to_excel(file_names, excel_path)
