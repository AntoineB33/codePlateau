import pandas as pd
import os


def convert_csv_to_xlsx(folder_path, xlsx_folder_path=""):
    # List all files in the given folder
    files = [file for file in os.listdir(folder_path) if file.endswith(".csv") or file.endswith(".CSV")]

    # Process each file
    for file in files:
        # Construct full file path
        file_path = os.path.join(folder_path, file)

        # Read the CSV file content
        with open(file_path, "r") as f:
            lines = f.readlines()

        # Check if the first line contains only integers
        first_line = lines[0].strip()
        if all(item.isdigit() for item in first_line.split(",")):
            # If the first line contains only integers, do not skip it
            data = "".join(lines)
        else:
            # If the first line contains non-integer values, skip it
            data = "".join(lines[1:])

        # Replace semicolons with commas in the data
        if "," not in data:
            data = data.replace(";", ",")

        # Write the updated content back to the CSV file (if modified)
        with open(file_path, "w") as f:
            f.write(data)

        # Read the CSV file into a DataFrame
        df = pd.read_csv(file_path)

        # Create a new Excel file path
        if not xlsx_folder_path:
            xlsx_folder_path = folder_path
        new_file_path = os.path.join(xlsx_folder_path, file.replace(".CSV", ".xlsx").replace(".csv", ".xlsx"))

        # Write data to an Excel file
        df.to_excel(new_file_path, index=False)
        print(f"Converted '{file}' to '{new_file_path}'")


# Specify the folder containing the CSV files
folder_path = r"C:\Users\comma\Documents\travail\Polytech\stage s8\gihtub\codePlateau\Resultats exp bag_couverts\Resultats exp bag_couverts\27_05_24"
xlsx_folder_path = r"C:\Users\comma\Documents\travail\Polytech\stage s8\gihtub\codePlateau\Resultats exp bag_couverts\Resultats exp bag_couverts\27_05_24_xlsx"
convert_csv_to_xlsx(folder_path, xlsx_folder_path)
