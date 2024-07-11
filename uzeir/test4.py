import os
import pandas as pd
import numpy as np
import math
from plotly.graph_objects import Figure, Scatter
from tkinter import filedialog, Tk
from scipy.signal import find_peaks, peak_prominences
import re

from scipy.fft import fft

import openpyxl
import xlwings as xw

import matplotlib.pyplot as plt
import time

import win32com.client as win32

# to normalize and reshape
from sklearn.preprocessing import MinMaxScaler
from tensorflow.keras.preprocessing.sequence import pad_sequences

# Build and Train the CNN Model
from tensorflow.keras.models import Sequential
from tensorflow.keras.layers import Conv1D, MaxPooling1D, Flatten, Dense
from tensorflow.keras.utils import to_categorical

""""///////////////////Variables globales///////////////////"""
poids_min = float("inf")
debut_ind = 0
fin_ind = 0
indice_debut = 0
indice_fin = 0
int_time = 0.2
plate_weight_min = 100
min_bite_duration = 1  # Minimum bite duration in seconds
min_bite_weight = 2
min_inactivity = 1
min_peak = 0
min_plate_weight = 700
noActivity_bgColor = 0xCDCDCC
activityWithBite_bgColor = 0xEBBE45
activityWithoutBite_bgColor = 0xCAECC7
module_name = "addData"  # Name for the new module if it needs to be added
fichier_names=[]
dataToExcel = []
recap_sheet_name = "Resultats_merged"
recap_titles = [
    "Bouchees",
    "Proportion_activite",
    "Duree_activite_min",
    "Duree_activite_max",
    "Duree_activite_mean",
    "Duree_activite_Totale",
    "Action",
    "Poids_Conso",
    "Poids_Moyen",
    "Poids_Min",
    "Poids_Max",
    "Sd_Poids",
    "Travail_Moyen",
    "Travail_Min",
    "Travail_Max",
    "Sd_Travail",
    "Force_Max",
    "Nb_segment_tot",
    "Nb_segment_bouchee",
    "Nb_segment_action_sans_bouchee",
    "Nb_segment_inacivite",
    "Duree_segm_bouchee_Moyenne",
    "Duree_segm_bouchee_Min",
    "Duree_segm_bouchee_Max",
    "Duree_segm_bouchee_Sd",
    "Duree_segm_action-b_Moyenne",
    "Duree_segm_action-b_Min",
    "Duree_segm_action-b_Max",
    "Duree_segm_action-b_Sd",
    "Duree_segm_inacivite_Moyenne",
    "Duree_segm_inacivite_Min",
    "Duree_segm_inacivite_Max",
    "Duree_segm_inacivite_Sd"
]
segm_titles = [
    "duree",
    "poids",
    "ecart de temps",
    "travail",
    "Force Max",
    "colors"
]
couverts = [
    "Baguettes",
    "Piquer",
    "Ramasser",
    "Couper"
]
"""/////////////////////////////////////////////////////////"""


def convert_csv_to_xlsx(folder_path, xlsx_folder_path=""):
    # List all files in the given folder
    files = [file for file in os.listdir("C:\\Users\\abarb\\Documents\\travail\\stage et4\\travail\\codePlateau\\data\\A envoyer_antoine(non corrompue)\\A envoyer\\Expériences plateaux") if file.endswith(".csv") or file.endswith(".CSV")]
    folder_path == "C:\\Users\\abarb\\Documents\\travail\\stage et4\\travail\\codePlateau\\data\\A envoyer_antoine(non corrompue)\\A envoyer\\Expériences plateaux"
    files = [file for file in os.listdir(folder_path) if file.endswith(".csv") or file.endswith(".CSV")]

    # Process each file
    for file in files:
        # Construct full file path
        file_path = os.path.join(folder_path, file)

        # Read the CSV file content
        with open(file_path, "r", encoding="ISO-8859-1") as f:
            lines = f.readlines()

        data = "".join(lines)
            
        # # Check if the first line contains only integers
        # first_line = lines[0].strip()
        # if all(item.isdigit() for item in first_line.split(",")):
        #     # If the first line contains only integers, do not skip it
        #     data = "".join(lines)
        # else:
        #     # If the first line contains non-integer values, skip it
        #     data = "".join(lines[1:])

        # Replace semicolons with commas in the data
        if "," not in data:
            data = data.replace(";", ",")
            with open(file_path, "w", encoding="ISO-8859-1") as f:
                f.write(data)

        # Read the CSV file into a DataFrame
        df = pd.read_csv(file_path, encoding="ISO-8859-1")

        # Ensure the output folder path exists
        if not xlsx_folder_path:
            xlsx_folder_path = folder_path
        os.makedirs(xlsx_folder_path, exist_ok=True)

        # Process each column against the first column as abscissa
        for col in df.columns[1:]:
            if not df[col].isnull().all():  # Ensure the column is not empty
                new_df = df[[df.columns[0], col]]

                # Create a new Excel file path
                new_file_path = os.path.join(
                    xlsx_folder_path, 
                    f"{os.path.splitext(file)[0]}_{col}.xlsx"
                )

                # Write data to an Excel file
                new_df.to_excel(new_file_path, index=False)
                # print(f"Converted '{file}' column '{col}' to '{new_file_path}'")

# Function to add or update the sheet if it doesn't exist
def add_sheet_if_not_exists2(wb, sheet_name):
    if sheet_name not in [sheet.name for sheet in wb.sheets]:
        wb.sheets.add(sheet_name)

# Function to update VBA module with code from .bas file
def update_vba_module_from_bas2(wb, bas_file_path):
    # Read the .bas file contents
    with open(bas_file_path, 'r') as file:
        vba_code = file.read()
    
    # Check if the module exists
    module_exists = False
    for component in wb.api.VBProject.VBComponents:
        if component.Name == module_name:
            module_exists = True
            module = component
            break

    # If the module doesn't exist, create it
    if not module_exists:
        module = wb.api.VBProject.VBComponents.Add(1)
        module.Name = module_name

    # Replace the existing code with the new code from the .bas file
    module.CodeModule.DeleteLines(1, module.CodeModule.CountOfLines)
    module.CodeModule.AddFromString(vba_code)

def rgb_to_bgr(rgb_color):
    red = (rgb_color >> 16) & 0xFF
    green = (rgb_color >> 8) & 0xFF
    blue = rgb_color & 0xFF
    bgr_color = (blue << 16) | (green << 8) | red
    return bgr_color

def ensure_sheets_exist(workbook, *sheet_names):
    for sheet_name in sheet_names:
        sheet_exists = False
        for sheet in workbook.Sheets:
            if sheet.Name == sheet_name:
                sheet_exists = True
                break
        if not sheet_exists:
            workbook.Sheets.Add().Name = sheet_name
            
def update_vba_module_from_bas(workbook, module_name, bas_file_path):
    with open(bas_file_path, 'r') as file:
        bas_code = file.read()

    vb_component = None
    for vb_comp in workbook.VBProject.VBComponents:
        if vb_comp.Name == module_name:
            vb_component = vb_comp
            break
    
    if vb_component is None:
        vb_component = workbook.VBProject.VBComponents.Add(1)
        vb_component.Name = module_name
    
    vb_component.CodeModule.DeleteLines(1, vb_component.CodeModule.CountOfLines)
    vb_component.CodeModule.AddFromString(bas_code)

def open_xlsm(full_file_path, module_name, *sheet_names):
    # def execute_vba_function(filename, vba_function_name, bas_file_path, module_name, open_all=False):
    # Initialize Excel application
    excel = win32.Dispatch('Excel.Application')

    # Check if the file exists, if not, create it
    if not os.path.exists(full_file_path):
        workbook = excel.Workbooks.Add()
        workbook.SaveAs(full_file_path, FileFormat=52)
        file_existed = False
    else:
        file_existed = True

    # Check if the file is already open
    workbook_open = False
    if file_existed:
        for wb in excel.Workbooks:
            if wb.FullName == full_file_path:
                workbook_open = True
                workbook = wb
                break

    # If the workbook is not already open, open it
    if not workbook_open:
        workbook = excel.Workbooks.Open(full_file_path)

    ensure_sheets_exist(workbook, *sheet_names)
        
    if sheet_names:
        # Update the VBA module with the code from the .bas file
        bas_file_path = "uzeir\\" + full_file_path.split("\\")[-1].replace('.xlsm', '.bas')
        update_vba_module_from_bas(workbook, module_name, bas_file_path)

    return excel, workbook, workbook_open

def open_xlsm2(full_file_path, open_all, *sheet_names):
    the_file_exists = True
    if not os.path.exists(full_file_path):
        app = xw.App(visible=open_all)
        time.sleep(1)
        app.books[0].close()
        the_file_exists = False
        wb = app.books.add()

    is_open = False
    if the_file_exists:
        # Check if the specific workbook is already open by its full path
        app = xw.apps.active
        if app:
            for book in xw.books:
                if book.fullname == full_file_path:
                    wb = book
                    is_open = True
                    break
        if not is_open:
            app = xw.App(visible=open_all)
            time.sleep(3)
            app.books[0].close()
            time.sleep(3)
            wb = app.books.open(full_file_path)
            time.sleep(3)


    # Add the sheet if it doesn't exist
    for sheet_name in sheet_names:
        if sheet_name!="colors":
            add_sheet_if_not_exists(wb, sheet_name)

    if sheet_names:
        # Update the VBA module with the code from the .bas file
        bas_file_path = "uzeir\\" + full_file_path.split("\\")[-1].replace('.xlsm', '.bas')
        update_vba_module_from_bas(wb, bas_file_path)
    return app, wb, the_file_exists, is_open


def close_xlsm(excel, workbook, workbook_open, open_all):
    # Save the workbook
    workbook.Save()

    # Close the workbook if it wasn't open before, unless open_all is True
    if not workbook_open and not open_all:
        workbook.Close()

    # Make Excel visible if open_all is True
    if open_all:
        excel.Visible = True

def close_xlsm2(app, wb, full_file_path, the_file_exists, is_open, open_all):
    # Save the workbook and close it if it was not already open
    if not the_file_exists:
        wb.save(full_file_path)
    else:
        wb.save()
    if not open_all and not is_open:
        wb.close()
        app.quit()

def calculate_statistics(numbers):
    # Calculate the mean
    mean = sum(numbers) / len(numbers)
    
    # Calculate the minimum
    min_value = min(numbers)
    
    # Calculate the maximum
    max_value = max(numbers)
    
    # Calculate the variance
    variance = sum((x - mean) ** 2 for x in numbers) / len(numbers)
    
    # Calculate the standard deviation
    standard_deviation = math.sqrt(variance)
    
    return {
        'mean': mean,
        'min': min_value,
        'max': max_value,
        'sd': standard_deviation
    }

def calculate_work(interval_df):
    # Create a copy of the 'Ptot' column to avoid modifying the original DataFrame
    Ptot_converted = interval_df['Ptot'] * 9.81 / 1000
    
    # Calculate the average weight (force) in the interval
    avg_weight = Ptot_converted.mean()
    
    # Calculate displacement using trapezoidal rule for integration
    # Here we assume constant velocity, displacement = velocity * time
    # For simplicity, assuming unit displacement for each time interval
    displacement = np.trapz(np.ones_like(interval_df['time']), interval_df['time'])
    
    # Work done
    work_done = avg_weight * displacement
    
    return work_done

def count_intervals_above_line(df, height, windows):
    above_line = df['Ptot'] > height
    intervals = []
    in_interval = False
    start = None

    for i, row in df.iterrows():
        if above_line[i]:
            if not in_interval:
                in_interval = True
                start = row['time']
        else:
            if in_interval:
                end = row['time']
                in_interval = False
                intervals.append((start, end))
    
    if in_interval:
        end = df.iloc[-1]['time']
        intervals.append((start, end))

    valid_intervals = []
    for interval in intervals:
        for window in windows:
            if interval[0] <= window[1] and interval[1] >= window[0]:
                valid_intervals.append(interval)
                break

    return len(valid_intervals), valid_intervals

def find_minimum_height(df, target_intervals, windows):
    unique_heights = sorted(df['Ptot'].unique())
    low, high = 0, len(unique_heights) - 1
    valid_intervals = []

    while low <= high:
        mid = (low + high) // 2
        height = unique_heights[mid]
        count, intervals = count_intervals_above_line(df, height, windows)

        if count == target_intervals:
            valid_intervals = intervals
            high = mid - 1
        elif count < target_intervals:
            high = mid - 1
        else:
            low = mid + 1

    return valid_intervals


def count_intervals_above_line(df, height, windows):
    above_line = df['Ptot'] > height
    intervals = []
    in_interval = False
    start = None

    for i, row in df.iterrows():
        if above_line[i]:
            if not in_interval:
                in_interval = True
                start = row['time']
        else:
            if in_interval:
                end = row['time']
                in_interval = False
                intervals.append((start, end))
    
    if in_interval:
        end = df.iloc[-1]['time']
        intervals.append((start, end))

    valid_intervals = []
    for interval in intervals:
        interval_start_index = df[df['time'] == interval[0]].index[0]
        interval_end_index = df[df['time'] == interval[1]].index[0]
        for window in windows:
            if interval_start_index <= window[1] and interval_end_index >= window[0]:
                valid_intervals.append(interval)
                break

    return len(valid_intervals), valid_intervals

def find_minimum_height(df, target_intervals, windows):
    min_height = df['Ptot'].min()
    max_height = df['Ptot'].max()
    current_height = min_height

    while current_height <= max_height:
        count, intervals = count_intervals_above_line(df, current_height, windows)
        if count == target_intervals:
            return current_height, intervals
        current_height += 1

    return None, []

def find_minimum_height(df, target_intervals, windows):
    unique_heights = sorted(df['Ptot'].unique())
    low, high = 0, len(unique_heights) - 1
    result = None
    valid_intervals = []

    while low <= high:
        mid = (low + high) // 2
        height = unique_heights[mid]
        count, intervals = count_intervals_above_line(df, height, windows)

        if count in (target_intervals, target_intervals + 1) :
            result = height
            valid_intervals = intervals
            high = mid - 1
        elif count < target_intervals:
            high = mid - 1
        else:
            low = mid + 1

    return result, valid_intervals

def find_minimum_height(df, target_intervals, windows):
    min_height = df['Ptot'].min()
    max_height = df['Ptot'].max()
    height = min_height
    result = None
    valid_intervals = []

    while height <= max_height:
        count, intervals = count_intervals_above_line(df, height, windows)

        if count in (target_intervals, target_intervals + 1):
            result = height
            if valid_intervals:
                break
            else:
                valid_intervals = intervals
                height += 10

        height += 1

    return result, valid_intervals

def find_minimum_height(df, target_intervals, windows):
    min_height = df['Ptot'].min()
    max_height = df['Ptot'].max()
    current_height = min_height
    num_points = len(df)
    threshold_points = num_points * 0.005

    prevCurrent_height, prevIntervals = 0, []

    while current_height <= max_height:
        count, intervals = count_intervals_above_line(df, current_height, windows)
        num_points_below = len(df[df['Ptot'] <= current_height])

        if count > target_intervals * 2 + 1 or num_points_below > threshold_points:
            return prevCurrent_height, prevIntervals
        prevCurrent_height, prevIntervals = current_height, intervals
        current_height += 1

    return None, []


def find_minimum_height(df, target_intervals, windows):
    unique_heights = sorted(df['Ptot'].unique())
    num_points = len(df)
    threshold_points = num_points * 0.005
    num_points_below = 0

    prevCurrent_height, prevIntervals = 0, []
    
    for height in unique_heights:
        count, intervals = count_intervals_above_line(df, height, windows)
        if count > target_intervals * 2 + 1 or num_points_below > threshold_points:
            return prevCurrent_height, prevIntervals
        prevCurrent_height, prevIntervals = height, intervals
        num_points_below += 1

    return None, []

def find_bites(dossier, dossier_graphique, date_folder, xlsm_recap, xlsm_recap_segments, file_PlateauExp = None, startCell_couverts = None, file = None, writeFileNames = False, to_update_excels = True, open_all = True, graph = True):
    global fichier_names

    if file_PlateauExp:
        # Open the workbook
        excel, workbook, workbook_open = open_xlsm(file_PlateauExp, module_name)
        
        sheet = workbook.Sheets(1)
        # start_range = sheet.Range(startCell_couverts)
        values = []
        # current_cell = start_range


        

        # Read cells from the start cell downwards until an empty cell is encountered
        row_number = sheet.Range(startCell_couverts).Row
        column_number = sheet.Range(startCell_couverts).Column
        while True:
            cell_value = sheet.Cells(row_number, column_number).Value
            if cell_value is None:
                break
            values.append(cell_value)
            row_number += 1
        
        # Close the workbook
        close_xlsm(excel, workbook, workbook_open, open_all)

        
        # Process the values
        processed_values = []
        for row, value in enumerate(values):
            # Split the value by "-" or "+"
            parts = re.split(r'[-+]', value)
            # Remove any trailing digits
            cleaned_parts = [re.sub(r'\d+$', '', part).strip() for part in parts]
            for part in cleaned_parts:
                if part not in couverts:
                    print(f"{part} n'est pas reconnu à la ligne {row + sheet.Range(startCell_couverts).Row}.")
                    return
            # Add the cleaned parts as a sublist
            processed_values.append(cleaned_parts)
    



    fichiers = []
    for f in os.listdir(dossier):
        if f.endswith(".xlsx") and (not file or f == file):
            fichiers.append(os.path.join(dossier, f))


    # Tableau_Final = pd.DataFrame(
    #     columns=[
    #         "Duree_Totale",
    #         "Poids_Conso",
    #         "Action",
    #         "Duree_activite_Totale",
    #         "Duree_activite_mean",
    #         "Duree_activite_max",
    #         "Duree_activite_min",
    #         "Proportion_activite_%",
    #         "Bouchees",
    #         "Num_fichier"
    #     ]
    # )

    print(dossier)
    recap_excel = dict()
    segment_excel = dict()


    # Fonction pour convertir le temps en minutes et secondes
    def convert_time(seconds):
        minutes = seconds // 60
        seconds = int(seconds) % 60
        return f"{minutes} min {seconds} s"


    def extract_features(segment):
        features = {}
        weights = segment["Ptot"].values
        times = segment["time"].values

        features["Duration"] = times[-1] - times[0]
        features["MaxWeight"] = np.max(weights)
        features["MinWeight"] = np.min(weights)
        features["MeanWeight"] = np.mean(weights)
        features["StdDevWeight"] = np.std(weights)

        # Calculate second derivative
        if len(weights) > 2:
            second_derivative = np.diff(np.diff(weights))
            features["SecondDerivative"] = np.mean(second_derivative)
        else:
            features["SecondDerivative"] = 0

        # Frequency components using FFT
        freq_components = np.abs(fft(weights))
        features["FrequencyComponent1"] = (
            freq_components[1] if len(freq_components) > 1 else 0
        )
        features["FrequencyComponent2"] = (
            freq_components[2] if len(freq_components) > 2 else 0
        )

        # Skewness and Kurtosis
        features["Skewness"] = pd.Series(weights).skew()
        features["Kurtosis"] = pd.Series(weights).kurtosis()

        # Peak analysis
        peaks, properties = find_peaks(weights, prominence=1)
        features["PeakCount"] = len(peaks)
        features["PeakProminence"] = (
            np.mean(properties["prominences"]) if "prominences" in properties else 0
        )

        return features
    
    fichier_names = []
    for fichier in fichiers:
        fichier_names.append(os.path.basename(fichier).split(".")[0])

    # Check if the graph folder exists, create it if not
    os.makedirs(dossier_graphique, exist_ok=True)
        
    valid_intervals = []
    
    # Traitement des fichiers
    for fileInd, fichier in enumerate(fichier_names):
        
        # recap_excel[fichier] = dict()
        # segment_excel[fichier] = dict()
        # for i in recap_titles:
        #     recap_excel[fichier][i] = 0
        # for i in segm_titles:
        #     segment_excel[fichier][i] = []
        # continue


        print(fichier)
        df = pd.read_excel(fichiers[fileInd], usecols=[0, 1])
        df.columns = ["time", "Ptot"]
        df["time"] = df["time"] / 1000
        df0 = df.copy()
        df = df[df["Ptot"] > 10]
        if not len(df):
            continue

        # df["Ptot"] = np.abs(df["Ptot"])
        # df = df[df["Ptot"] < 3000]

        # Filter smoothing time
        filtered_df = df.copy()
        indices_time = [0]

        for i in range(1, len(filtered_df)):
            if filtered_df["time"].iloc[i] - filtered_df["time"].iloc[indices_time[-1]] >= int_time:
                indices_time.append(i)

        filtered_df = filtered_df.iloc[indices_time].reset_index(drop=True)

        if filtered_df.empty:
            continue  # Skip further processing if no data remains after filtering

        # Filter smoothing weight
        i = 0

        while i < len(filtered_df) - 1:
            val_ini = filtered_df["Ptot"].iloc[i]
            j = i + 1

            while j < len(filtered_df) and np.abs(filtered_df["Ptot"].iloc[j] - val_ini) < min_bite_weight:
                j += 1

            if j <= len(filtered_df):
                filtered_df.loc[i : j - 1, "Ptot"] = val_ini
            i = j

        df = filtered_df  # Update df with the noise-filtered data

        # Calcul du temps d'activité et du nombre de bouchée
        activity_time = 0
        peaks_x = []
        peaks_y = []

        # Peak detection
        peaks, _ = find_peaks(
            df["Ptot"], height=100, distance=1
        )  # Reduced distance for more sensitivity
        prominences = peak_prominences(df["Ptot"], peaks)[0]

        # Filter peaks based on prominence
        poids_seuil = 10  # Lowered threshold for more sensitivity
        significant_peaks_indices = np.where(prominences > poids_seuil)[0]
        significant_peaks = peaks[significant_peaks_indices]
        significant_prominences = prominences[significant_peaks_indices]

        # Further filter peaks to ensure sufficient weight difference
        weight_diff_threshold = 2  # Lowered threshold for more sensitivity
        valid_peaks_indices = []
        for idx in significant_peaks:
            if (
                (idx - 1 >= 0 and idx + 1 < len(df))
                and (
                    df["Ptot"].iloc[idx] - df["Ptot"].iloc[idx - 1] > weight_diff_threshold
                )
                and (
                    df["Ptot"].iloc[idx] - df["Ptot"].iloc[idx + 1] > weight_diff_threshold
                )
            ):
                # Check if weight changes significantly before and after the peak
                window = 5  # Define a window around the peak to check weight change
                if (
                    (idx - window >= 0 and idx + window < len(df))
                    and (
                        abs(df["Ptot"].iloc[idx - window] - df["Ptot"].iloc[idx])
                        > weight_diff_threshold
                    )
                    and (
                        abs(df["Ptot"].iloc[idx + window] - df["Ptot"].iloc[idx])
                        > weight_diff_threshold
                    )
                ):
                    valid_peaks_indices.append(idx)

        valid_peaks = np.array(valid_peaks_indices)
        valid_prominences = peak_prominences(df["Ptot"], valid_peaks)[0]

        if len(valid_peaks) == 0:
            continue

        # Filter closely spaced peaks and merge windows
        min_diff = 50  # Minimum difference in indices between consecutive peaks

        i = valid_peaks[-1]
        while i + 1 < len(df):
            if df["Ptot"].iloc[i] != df["Ptot"].iloc[i + 1]:
                valid_peaks = np.append(valid_peaks, i + 1)
                valid_prominences = np.append(valid_prominences, 0)
            i+=1
        allPeaksFound = False
        while not allPeaksFound:
            stop_the_bite = True
            allPeaksFound = True
            add_peak_update_next = True
            activity_time = 0
            final_peaks_indices = []
            merged_windows = []
            is_bite = []
            associatedWith = []
            window_start, window_end = 0, 0
            Duree_activite_min = float("inf")
            Duree_activite_max = 0
            bouchees = 0
            i = 0
            while i <= len(valid_peaks):
                if final_peaks_indices:
                    lastI = i == len(valid_peaks)
                    if stop_the_bite:
                        # increase the window to the left
                        exploring_horizontal = 5
                        for j in range(window_start, 0, -1):
                            if df["Ptot"].iloc[j] == df["Ptot"].iloc[j - 1]:
                                exploring_horizontal -= 1
                                if exploring_horizontal == 0:
                                    break
                            else:
                                exploring_horizontal = 5
                                window_start = j - 1
                    upTo = len(df) - 1 if lastI else valid_peaks[i]
                    stop_the_bite = lastI or (upTo - final_peaks_indices[-1]) > min_diff
                    if stop_the_bite:
                        # increase the window to the right
                        exploring_horizontal = 5
                        for j in range(window_end, upTo):
                            if df["Ptot"].iloc[j] == df["Ptot"].iloc[j + 1]:
                                exploring_horizontal -= 1
                                if exploring_horizontal == 0:
                                    break
                            else:
                                exploring_horizontal = 5
                                window_end = j + 1
                        stop_the_bite = lastI or (upTo - window_end) > min_diff
                    if not stop_the_bite:
                        # cut into two bites if a long period of inactivity is found in the bite window
                        last_activity = window_end
                        for window_endi in range(window_end, upTo):
                            if (
                                df["Ptot"].iloc[window_endi]
                                != df["Ptot"].iloc[window_endi + 1]
                            ):
                                if (
                                    df["time"].iloc[window_endi]
                                    - df["time"].iloc[last_activity]
                                    > min_inactivity
                                ):
                                    stop_the_bite = True
                                    window_end = last_activity
                                    break
                                last_activity = window_endi + 1
                    if stop_the_bite:
                        if window_start == 0 and df0["time"].iloc[0]!=df["time"].iloc[0]:
                            del final_peaks_indices[0]
                        else:
                            if window_end == len(df)-1 and df0["time"].iloc[len(df0) - 1]!=df["time"].iloc[len(df)-1]:
                                del final_peaks_indices[-1]
                                break
                            merged_windows.append((window_start, window_end))
                            Duree_activity = (
                                df["time"].iloc[window_end] - df["time"].iloc[window_start]
                            )
                            activity_time += Duree_activity
                            if Duree_activite_min > Duree_activity:
                                Duree_activite_min = Duree_activity
                            if Duree_activite_max < Duree_activity:
                                Duree_activite_max = Duree_activity
                            # check if the activity decreases the amount of food
                            diff = df["Ptot"].iloc[window_end] - df["Ptot"].iloc[window_start]
                            associatedWith.append(-1)
                            if diff <= -min_bite_weight:
                                for index, prev_diff in enumerate(is_bite):
                                    if abs(prev_diff + diff) < prev_diff / 20 + 1:
                                        diff = 0
                                        is_bite[index] = 0
                                        associatedWith[-1] = index
                                        break
                                if diff:
                                    bouchees += 1
                                    fin_ind = len(merged_windows)
                            # check if the food quantity has decreased before this bite
                            if (
                                len(merged_windows) > 1
                                and df["Ptot"].iloc[window_start]
                                != df["Ptot"].iloc[merged_windows[-2][1]]
                                and 10):
                                # go back to see where would be the missing bite
                                # last_quantity = df["Ptot"].iloc[merged_windows[-2][1]]
                                in_peak = False
                                for j in range(merged_windows[-2][1] + 1, window_start):
                                    # if df["time"].iloc[j]>=390.95:
                                    #     print(5)
                                    if j>0 and df["Ptot"].iloc[j] != df["Ptot"].iloc[j - 1]:
                                        # if j == 719 or j == 722 or j == 723 or j == 725:
                                        #     print(7)
                                        valid_peaks = np.insert(valid_peaks, firstPeakAfterPrevAct, j)
                                        firstPeakAfterPrevAct += 1
                                        valid_prominences = np.insert(
                                            valid_prominences, i - 1, 0
                                        )
                                        allPeaksFound = False
                                        i += 1
                                    # if df["Ptot"].iloc[j] > last_quantity + min_peak:
                                    #     valid_peaks = np.insert(valid_peaks, i - 1, j)
                                    #     valid_prominences = np.insert(
                                    #         valid_prominences, i - 1, 0
                                    #     )
                                    #     allPeaksFound = False
                                    #     i += 1
                                    #     in_peak = True
                                    #     exploring_horizontal = 5
                                    # elif (
                                    #     not in_peak
                                    #     and df["Ptot"].iloc[j] < last_quantity
                                    #     and j > 0
                                    #     and df["Ptot"].iloc[j] == df["Ptot"].iloc[j - 1]
                                    # ):
                                    #     last_quantity = df["Ptot"].iloc[j]
                                    # elif j>0 and df["Ptot"].iloc[j] == df["Ptot"].iloc[j - 1]:
                                    #     if exploring_horizontal == 0:
                                    #         last_quantity = df["Ptot"].iloc[j]
                                    #     exploring_horizontal-=1
                                    #     in_peak = False
                            is_bite.append(diff)
                            firstPeakAfterPrevAct = i
                        add_peak_update_next = not lastI
                    else:
                        # Merge peaks if they are close
                        window_end = valid_peaks[i]
                        if (
                            valid_prominences[i]
                            > valid_prominences[
                                np.where(valid_peaks == final_peaks_indices[-1])[0][0]
                            ]
                        ):
                            final_peaks_indices[-1] = valid_peaks[i]
                if add_peak_update_next:
                    final_peaks_indices.append(valid_peaks[i])
                    window_start = valid_peaks[i]
                    window_end = valid_peaks[i]
                    add_peak_update_next = False
                i += 1


        # Calcul du temps d'activité
        # if len(significant_peaks_x) > 0:
        #     activity_start_time = significant_peaks_x[0]
        #     activity_end_time = significant_peaks_x[-1]
        #     activity_time = math.trunc(activity_end_time - activity_start_time)
        # else:
        #     activity_time = 0


        if bouchees:
            # debut_ind = 0
            # while -min_bite_weight < is_bite[debut_ind]:
            #     debut_ind += 1
            # merged_windows = merged_windows[debut_ind:fin_ind]
            # final_peaks_indices = final_peaks_indices[debut_ind:fin_ind]
            # is_bite = is_bite[debut_ind:fin_ind]

            debut_time = merged_windows[0][0]
            fin_time = merged_windows[-1][1]
            poids_consome = df["Ptot"].iloc[debut_time] - df["Ptot"].iloc[fin_time]
            temps_repas = df["time"].iloc[fin_time] - df["time"].iloc[debut_time]
        else:
            poids_consome = 0
            temps_repas = 0
        significant_peaks_x = df["time"].iloc[final_peaks_indices].values
        significant_peaks_y = df["Ptot"].iloc[final_peaks_indices].values
        ratio = activity_time / temps_repas if temps_repas > 0 else 0

        print(f"Le poids consommé pendant le repas est : {poids_consome}")
        print(f"La durée du repas est : {convert_time(temps_repas)}")
        print(f"Le temps d'activité est : {convert_time(activity_time)}")
        print(f"Le ratio d'activité est : {ratio}")
        print(f"Le nombre de bouchée pendant le repas est : {bouchees}")

        # for couvert in processed_values[fileInd]:
        #     fichier += "_" + couvert
        recap_excel[fichier] = dict()
        recap_excel[fichier]["Bouchees"] = bouchees
        recap_excel[fichier]["Proportion_activite"] = round(ratio * 100, 1)
        recap_excel[fichier]["Duree_activite_min"] = Duree_activite_min
        recap_excel[fichier]["Duree_activite_max"] = Duree_activite_max
        recap_excel[fichier]["Duree_activite_mean"] = round(activity_time / bouchees, 3)
        recap_excel[fichier]["Duree_activite_Totale"] = activity_time
        recap_excel[fichier]["Action"] = len(merged_windows)
        recap_excel[fichier]["Poids_Conso"] = poids_consome

        
        segment_excel[fichier] = dict()

        # storing the duration of each action
        for i in segm_titles:
            segment_excel[fichier][i] = []
        poids_bouchees = []
        travail_bouchees = []
        force_max = []
        duree_bouchees = []
        duree_activitee_sans_b = []
        duree_inactivitee = []
        for index, window in enumerate(merged_windows):
            segment_excel[fichier]["duree"].append(df["time"].iloc[window[1]] - df["time"].iloc[window[0]])
            segment_excel[fichier]["ecart de temps"].append(df["time"].iloc[window[1]] - df["time"].iloc[window[0]])
            interval = df.iloc[window[0]:window[1]+1]
            travail_bouchees.append(calculate_work(interval))
            segment_excel[fichier]["travail"].append(travail_bouchees[-1])
            force_max.append(interval["Ptot"].max() - df["Ptot"].iloc[window[0]])
            segment_excel[fichier]["Force Max"].append(force_max[-1])
            segment_excel[fichier]["colors"].append(int(is_bite[index] <= -min_bite_weight))
            if segment_excel[fichier]["colors"][-1]:
                segment_excel[fichier]["poids"].append(-is_bite[index])
                poids_bouchees.append(-is_bite[index])
                duree_bouchees.append(segment_excel[fichier]["duree"][-1])
            else:
                segment_excel[fichier]["poids"].append(0)
                duree_activitee_sans_b.append(segment_excel[fichier]["duree"][-1])
            if index + 1 < len(merged_windows):
                diff = df["time"].iloc[merged_windows[index + 1][0]] - df["time"].iloc[window[1]]
                if diff:
                    segment_excel[fichier]["duree"].append(diff)
                    segment_excel[fichier]["poids"].append(0)
                    segment_excel[fichier]["ecart de temps"].append(0)
                    segment_excel[fichier]["travail"].append(0)
                    segment_excel[fichier]["Force Max"].append(0)
                    segment_excel[fichier]["colors"].append(2)
                    duree_inactivitee.append(diff)
                
        stats = calculate_statistics(poids_bouchees)
        recap_excel[fichier]["Poids_Moyen"] = stats["mean"]
        recap_excel[fichier]["Poids_Min"] = stats["min"]
        recap_excel[fichier]["Poids_Max"] = stats["max"]
        recap_excel[fichier]["Sd_Poids"] = stats["sd"]
        stats = calculate_statistics(travail_bouchees)
        recap_excel[fichier]["Travail_Moyen"] = stats["mean"]
        recap_excel[fichier]["Travail_Min"] = stats["min"]
        recap_excel[fichier]["Travail_Max"] = stats["max"]
        recap_excel[fichier]["Sd_Travail"] = stats["sd"]
        recap_excel[fichier]["Force_Max"] = max(force_max)
        recap_excel[fichier]["Nb_segment_tot"] = len(merged_windows)
        recap_excel[fichier]["Nb_segment_bouchee"] = len(duree_bouchees)
        recap_excel[fichier]["Nb_segment_action_sans_bouchee"] = len(duree_activitee_sans_b)
        recap_excel[fichier]["Nb_segment_inacivite"] = len(duree_inactivitee)
        stats = calculate_statistics(duree_bouchees)
        recap_excel[fichier]["Duree_segm_bouchee_Moyenne"] = stats["mean"]
        recap_excel[fichier]["Duree_segm_bouchee_Min"] = stats["min"]
        recap_excel[fichier]["Duree_segm_bouchee_Max"] = stats["max"]
        recap_excel[fichier]["Duree_segm_bouchee_Sd"] = stats["sd"]
        stats = calculate_statistics(duree_activitee_sans_b)
        recap_excel[fichier]["Duree_segm_action-b_Moyenne"] = stats["mean"]
        recap_excel[fichier]["Duree_segm_action-b_Min"] = stats["min"]
        recap_excel[fichier]["Duree_segm_action-b_Max"] = stats["max"]
        recap_excel[fichier]["Duree_segm_action-b_Sd"] = stats["sd"]
        stats = calculate_statistics(duree_inactivitee)
        recap_excel[fichier]["Duree_segm_inacivite_Moyenne"] = stats["mean"]
        recap_excel[fichier]["Duree_segm_inacivite_Min"] = stats["min"]
        recap_excel[fichier]["Duree_segm_inacivite_Max"] = stats["max"]
        recap_excel[fichier]["Duree_segm_inacivite_Sd"] = stats["sd"]

        if file_PlateauExp:
            min_height, valid_intervals = find_minimum_height(df, len(processed_values[fileInd]), merged_windows)

            segments = []
            labels = []
            i = 0
            j = 0
            for startCouv, endCouv in valid_intervals:
                merged_windowsI = merged_windows[i:]
                for start, end in merged_windowsI:
                    if df["time"].iloc[end] > endCouv:
                        j += 1
                        break
                    segments.append(df['Ptot'][start:end+1].values)
                    labels.append(couverts.index(processed_values[fileInd][j]))
                    i+=1
            
            with open(r"C:\Users\abarb\Documents\travail\stage et4\travail\codePlateau\uzeir\interval_label.txt", 'a') as file:  # 'a' mode for appending
                for segment, label in zip(segments, labels):
                    segment_str = ','.join(map(str, segment))
                    file.write(f'{segment_str};{label}\n')

        if not graph:
            continue

        # Création de graphiques avec Plotly
        fig = Figure()
        fig.update_layout(
            title_text="Poids Total en fonction du temps",
            xaxis_title="Temps en secondes",
            yaxis_title="Poids en grammes",
        )
        fig.add_trace(Scatter(y=df["Ptot"], x=df["time"], mode="lines", name="Poids Total"))

        # Ajouter les fenêtres en rouge pour chaque pic
        for index, (start_idx, end_idx) in enumerate(merged_windows):
            # # Début de la fenêtre : recherche en arrière jusqu'à ce que le poids commence à augmenter
            # while (
            #     start_idx > 0
            #     and df["Ptot"].iloc[start_idx - 1] < df["Ptot"].iloc[start_idx]
            # ):
            #     start_idx -= 1

            # # Fin de la fenêtre : recherche en avant jusqu'à ce que le poids redevienne constant
            # while (
            #     end_idx < len(df) - 1
            #     and df["Ptot"].iloc[end_idx + 1] < df["Ptot"].iloc[end_idx]
            # ):
            #     end_idx += 1
            color = "Red" if is_bite[index] <= -min_bite_weight else "Gray"
            fig.add_shape(
                type="rect",
                x0=df["time"].iloc[start_idx],
                y0=df["Ptot"].min(),
                x1=df["time"].iloc[end_idx],
                y1=df["Ptot"].max(),
                line=dict(color=color, width=2),
                fillcolor=color,
                opacity=0.2,
            )
            if associatedWith[index] != -1:
                decalageVert = 200
                yPos = df["Ptot"].iloc[start_idx:end_idx].max() + decalageVert
                if yPos > df["Ptot"].max():
                    yPos = df["Ptot"].iloc[start_idx:end_idx].min() - decalageVert
                # Add the text annotation
                fig.add_annotation(
                    x=df["time"].iloc[(start_idx + end_idx) // 2],
                    y=yPos,
                    text=f"associated with action<br>starting at {df['time'].iloc[merged_windows[associatedWith[index]][0]]}",
                    showarrow=False,
                    font=dict(color="black", size=12),
                    bgcolor="white",
                    opacity=0.8,
                )
        for couvertInd, couvert in enumerate(valid_intervals):
            if couvertInd:
                fig.add_shape(
                    type="rect",
                    x0=df["time"].iloc[valid_intervals[couvertInd - 1]],
                    y0=df["Ptot"].min(),
                    x1=df["time"].iloc[couvert[0]],
                    y1=df["Ptot"].max(),
                    line=dict(color=color, width=2),
                    fillcolor="green",
                    opacity=0.2,
                )


        # fig.add_trace(
        #     Scatter(
        #         y=significant_peaks_y,
        #         x=significant_peaks_x,
        #         mode="markers",
        #         name="Pics détectés",
        #         marker=dict(color="red", size=8),
        #     )
        # )
        fig.add_annotation(
            x=df["time"].iloc[-1],
            y=df["Ptot"].max(),
            text=f"Poids consommé: {poids_consome} g<br>Durée du repas: {convert_time(temps_repas)}<br>Temps d'activité sur l'assiette: {convert_time(activity_time)}<br>Ratio d'activité: {ratio:.2f}<br>Nombre de bouchées: {bouchees}",
            showarrow=False,
            align="left",
            xanchor="right",
            yanchor="top",
        )

        # valid_peaks_x = df["time"].iloc[valid_peaks].values
        # valid_peaks_y = df["Ptot"].iloc[valid_peaks].values
        # fig.add_trace(
        #     Scatter(
        #         y=valid_peaks_y,
        #         x=valid_peaks_x,
        #         mode="markers",
        #         name="Pics significatifs",
        #         marker=dict(color="blue", size=8),
        #     )
        # )

        # Enregistrement des graphiques
        filepath = os.path.join(
            dossier_graphique,
            "Graph_{}.html".format(fichier_names[fileInd]),
        )

        fig.write_html(filepath)

        # Extract features for each bite and store in a list
        feature_list = []
        for start_idx, end_idx in merged_windows:
            segment = df.iloc[start_idx : end_idx + 1]
            features = extract_features(segment, )
            features["BiteID"] = len(feature_list) + 1  # Unique identifier for each bite
            # Add the known label here if available, e.g., features['Label'] = 'Fork'
            feature_list.append(features)

        # Convert to DataFrame
        features_df = pd.DataFrame(feature_list)

        # Save to CSV
        features_df.to_csv("bite_features.csv", index=False)

    if to_update_excels:
        # new_fichier_names = []
        # for fileInd, fichier in enumerate(fichier_names):
        #     for couvert in processed_values[fileInd]:
        #         new_fichier_names.append(fichier + "_" + couvert)
        # fichier_names = new_fichier_names
        fichier_names_rows = [name + date_folder for name in fichier_names]


        fills = [str(rgb_to_bgr(color)) for color in [activityWithoutBite_bgColor, activityWithBite_bgColor, noActivity_bgColor]]

            
        data_lst = []
        data_lst_segments = []
        for fileInd, fichier in enumerate(fichier_names_rows):
            if fichier_names[fileInd] in recap_excel:
                data_lst.append([fichier])
                data_lst_segments.append([])
                # loop throught the keys of new_excel
                for key in recap_titles:
                    data_lst[-1].append(str(recap_excel[fichier_names[fileInd]][key]))
                for key in segm_titles:
                    data_lst_segments[-1].append([])
                    for window in segment_excel[fichier_names[fileInd]][key]:
                        if key == "colors":
                            data_lst_segments[-1][-1].append(fills[window])
                        else:
                            data_lst_segments[-1][-1].append(str(round(window, 1)))
        
        excel, workbook, workbook_open = open_xlsm(xlsm_recap, module_name, recap_sheet_name)
        if writeFileNames:
            excel.Application.Run(f'{workbook.Name}!allFileName', recap_sheet_name, fichier_names_rows)
        row_found = excel.Application.Run(f'{workbook.Name}!SearchAndImportData', recap_sheet_name, "A", recap_titles, data_lst)
        close_xlsm(excel, workbook, workbook_open, open_all)
        # excel, workbook, workbook_open = open_xlsm(xlsm_recap_segments, module_name, *segm_titles)
        # if row_found:
        #     data_lst_segments.append([])
        #     wb.macro(module_name + ".ImportSegments")(row_found, segm_titles, data_lst_segments)
        # else:
        #     print(f"File name {fichier} not found in the main excel.")
        # close_xlsm(excel, workbook, workbook_open, open_all)

    # def analyseAI():



def train_AI(filename=r"C:\Users\abarb\Documents\travail\stage et4\travail\codePlateau\uzeir\interval_label.txt"):
    segments = []
    labels = []
    with open(filename, 'r') as file:
        for line in file:
            segment_str, label_str = line.strip().split(';')
            segment = list(map(float, segment_str.split(',')))
            label = int(label_str)
            segments.append(segment)
            labels.append(label)
                    
            
    # Normalize segments
    scaler = MinMaxScaler()

    normalized_segments = [scaler.fit_transform(segment.reshape(-1, 1)).flatten() for segment in segments]

    # Pad sequences to ensure equal length
    padded_segments = pad_sequences(normalized_segments, padding='post', dtype='float32')

    # Convert to numpy array and reshape for CNN
    X = np.array(padded_segments)
    y = np.array(labels)

    X = X.reshape((X.shape[0], X.shape[1], 1))



    # One-hot encode labels
    y = to_categorical(y, num_classes=3)

    # Build CNN model
    model = Sequential()
    model.add(Conv1D(filters=64, kernel_size=3, activation='relu', input_shape=(X.shape[1], 1)))
    model.add(MaxPooling1D(pool_size=2))
    model.add(Flatten())
    model.add(Dense(100, activation='relu'))
    model.add(Dense(3, activation='softmax'))

    # Compile the model
    model.compile(optimizer='adam', loss='categorical_crossentropy', metrics=['accuracy'])

    # Train the model
    model.fit(X, y, epochs=10, verbose=1)



    # Evaluate the model on training data (use validation data in practice)
    loss, accuracy = model.evaluate(X, y, verbose=0)
    print(f'Accuracy: {accuracy*100:.2f}%')



excel_all_path = r".\data\Resultats exp bag_couverts\Resultats exp bag_couverts\Tableau récapitulatif - new algo"
excel_segments_path = r".\data\Resultats exp bag_couverts\Resultats exp bag_couverts\durée_segments"


dossier = r".\data\Resultats exp bag_couverts\Resultats exp bag_couverts\27_05_24_xlsx"
# dossier = r".\data_du_bureau\xlsx"
# dossier = r".\filtered_data"

dossier_graphique = r".\data\Resultats exp bag_couverts\Resultats exp bag_couverts\27_05_24_graph"

date_folder = "_27_05_24"

# find_bites(dossier, dossier_graphique, date_folder,"2Plateaux-P1-couv")
# find_bites(dossier, dossier_graphique, date_folder)

dossier = r".\data\Resultats exp bag_couverts\Resultats exp bag_couverts\28_05_24_xlsx"
# dossier = r".\data_du_bureau\xlsx"
# dossier = r".\filtered_data"

dossier_graphique = r".\data\Resultats exp bag_couverts\Resultats exp bag_couverts\28_05_24_graph"

date_folder = "_28_05_24"

# find_bites(dossier, dossier_graphique, date_folder)















# path = r"C:\Users\abarb\Documents\travail\stage et4\travail\codePlateau\data\A envoyer(pate a modeler)\A envoyer"
path = r"C:\Users\abarb\Documents\travail\stage et4\travail\codePlateau\data\A envoyer_antoine(non corrompue)\A envoyer"
date_folder = ""

path += "\\"
dossier = path + "xlsx"
dossier_graphique = path + "graph"
dossier_recap = path + r"recap\recap.xlsm"
dossier_recap_segments = path + r"recap\segments.xlsm"


convert_csv_to_xlsx(path + "Expériences plateaux", dossier)

file_PlateauExp = "Plateaux_Exp.xlsx"
file_PlateauExp = r"C:\Users\abarb\Documents\travail\stage et4\travail\codePlateau\data\A envoyer_antoine(non corrompue)\A envoyer\Plateaux_Exp.xlsx"
startCell_couverts = "F4"
startCell_couverts = "E4"


# find_bites(dossier, dossier_graphique, date_folder, dossier_recap, dossier_recap_segments, file_PlateauExp, startCell_couverts, file = "180624_Dorian_Laura_P1.xlsx", writeFileNames = True)
find_bites(dossier, dossier_graphique, date_folder, dossier_recap, dossier_recap_segments, file_PlateauExp, startCell_couverts, file = "18_06_24_Benjamin_Roxane_P1.xlsx", writeFileNames = False, graph = False)
# find_bites(dossier, dossier_graphique, date_folder, dossier_recap, dossier_recap_segments, file_PlateauExp, startCell_couverts, writeFileNames = True, graph = False)
# find_bites(dossier, dossier_graphique, date_folder, "14_05_Benjamin.xlsx")


train_AI()