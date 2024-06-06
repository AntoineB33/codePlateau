import pandas as pd
import os


def convert_csv_to_xlsx(folder_path):
    # List all files in the given folder
    files = [file for file in os.listdir(folder_path) if file.endswith(".CSV")]

    # Process each file
    for file in files:
        # Construct full file path
        file_path = os.path.join(folder_path, file)
        # Read the CSV file
        df = pd.read_csv(file_path)
        # Create a new Excel file path
        new_file_path = os.path.join(folder_path, file.replace(".CSV", ".xlsx"))
        # Write data to an Excel file
        df.to_excel(new_file_path, index=False)
        print(f"Converted '{file}' to '{new_file_path}'")


# Specify the folder containing the CSV files
folder_path = r"C:\Users\comma\Documents\travail\Polytech\stage s8\gihtub\codePlateau"
convert_csv_to_xlsx(folder_path)
