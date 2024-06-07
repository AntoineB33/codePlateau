import pandas as pd
import os


def convert_csv_to_xlsx(folder_path, xlsx_folder_path = ""):
    # List all files in the given folder
    files = [file for file in os.listdir(folder_path) if file.endswith(".csv") or file.endswith(".CSV")]

    # Process each file
    for file in files:
        # Construct full file path
        file_path = os.path.join(folder_path, file)

        # Read the CSV file content
        with open(file_path, "r") as f:
            content = f.read()

        # Check if there are any commas in the content
        if "," not in content:
            # Replace semicolons with commas
            content = content.replace(";", ",")

            # Write the updated content back to the CSV file
            with open(file_path, "w") as f:
                f.write(content)

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
folder_path = r"data_du_bureau\csv"
xlsx_folder_path = r"data_du_bureau\xlsx"
convert_csv_to_xlsx(folder_path, xlsx_folder_path)
