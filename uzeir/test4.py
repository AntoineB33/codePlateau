import os
import pandas as pd
import numpy as np
import math
from plotly.graph_objects import Figure, Scatter
from tkinter import filedialog, Tk
from scipy.signal import find_peaks, peak_prominences
import re

import win32com.client
import win32api

from scipy.fft import fft

import os
import openpyxl
import xlwings as xw

import matplotlib.pyplot as plt
import time

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

import os
import pandas as pd

def convert_csv_to_xlsx(folder_path, xlsx_folder_path=""):
    # List all files in the given folder
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
def add_sheet_if_not_exists(wb, sheet_name):
    if sheet_name not in [sheet.name for sheet in wb.sheets]:
        wb.sheets.add(sheet_name)

# Function to update VBA module with code from .bas file
def update_vba_module_from_bas(wb, bas_file_path):
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

def open_xlsm(full_file_path, open_all, *sheet_names):
    the_file_exists = True
    if not os.path.exists(full_file_path):
        app = xw.App(visible=open_all)
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
            app.books[0].close()
            wb = app.books.open(full_file_path)


    # Add the sheet if it doesn't exist
    for sheet_name in sheet_names:
        if sheet_name!="colors":
            add_sheet_if_not_exists(wb, sheet_name)

    # Update the VBA module with the code from the .bas file
    bas_file_path = "uzeir\\" + full_file_path.split("\\")[-1].replace('.xlsm', '.bas')
    update_vba_module_from_bas(wb, bas_file_path)
    return app, wb, the_file_exists, is_open

def close_xlsm(app, wb, full_file_path, the_file_exists, is_open, open_all):
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
    
    # Convert weight from grams to Newtons
    interval_df['Ptot'] = interval_df['Ptot'] * 9.81 / 1000
    
    # Calculate the average weight (force) in the interval
    avg_weight = interval_df['Ptot'].mean()
    
    # Calculate displacement using trapezoidal rule for integration
    # Here we assume constant velocity, displacement = velocity * time
    # For simplicity, assuming unit displacement for each time interval
    displacement = np.trapz(np.ones_like(interval_df['time']), interval_df['time'])
    
    # Work done
    work_done = avg_weight * displacement
    
    return work_done

def find_bites(dossier, dossier_graphique, date_folder, xlsm_recap, xlsm_recap_segments, file_PlateauExp, startCell_couverts, file = None, writeFileNames = False, to_update_excels = True, open_all = True):
    global fichier_names

    
    # Open the workbook
    app, wb, the_file_exists, is_open = open_xlsm(file_PlateauExp, open_all)
    # Select the active sheet
    sheet = wb.sheets[0]
    # Get the row and column indices from the start cell
    start_row = sheet.range(startCell_couverts).row
    start_col = sheet.range(startCell_couverts).column
    
    # Read values from the start cell downwards
    values = []
    row = start_row
    while True:
        cell_value = sheet.range((row, start_col)).value
        if cell_value is None:
            break
        values.append(cell_value)
        row += 1
    
    # Close the workbook
    close_xlsm(app, wb, file_PlateauExp, the_file_exists, is_open, open_all)

    
    # Process the values
    processed_values = []
    for value in values:
        # Split the value by "-" or "+"
        parts = re.split(r'[-+]', value)
        # Remove any trailing digits
        cleaned_parts = [re.sub(r'\d+$', '', part) for part in parts]
        for part in cleaned_parts:
            if part not in couverts:
                print(f"{part} n'est pas reconnu à la ligne {row + start_row}")
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
        
    # Traitement des fichiers
    for fileInd, fichier in enumerate(fichier_names):
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
            if segment_excel[fichier]["colors"][index]:
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

        fig.add_trace(
            Scatter(
                y=significant_peaks_y,
                x=significant_peaks_x,
                mode="markers",
                name="Pics détectés",
                marker=dict(color="red", size=8),
            )
        )
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
        
        app, wb, the_file_exists, is_open = open_xlsm(xlsm_recap, open_all, recap_sheet_name)



        if writeFileNames:
            wb.macro(module_name + ".allFileName")(recap_sheet_name, fichier_names_rows)

            
        data_lst = []
        data_lst_segments = []
        for fileInd, fichier in enumerate(fichier_names_rows):
            data_lst.append([fichier])
            data_lst_segments.append([])
            # loop throught the keys of new_excel
            for key in recap_titles:
                data_lst[fileInd].append(str(recap_excel[fichier_names[fileInd]][key]))
            for key in segm_titles:
                data_lst_segments[fileInd].append([])
                for window in segment_excel[fichier_names[fileInd]][key]:
                    if key == "colors":
                        data_lst_segments[fileInd][-1].append(fills[window])
                    else:
                        data_lst_segments[fileInd][-1].append(str(round(window, 1)))
        row_found = wb.macro(module_name + ".SearchAndImportData")(recap_sheet_name, "A", recap_titles, data_lst)
        close_xlsm(app, wb, xlsm_recap, the_file_exists, is_open, open_all)
        # app, wb, the_file_exists, is_open = open_xlsm(xlsm_recap_segments, open_all, *segm_titles)
        # if row_found:
        #     data_lst_segments.append([])
        #     wb.macro(module_name + ".ImportSegments")(row_found, segm_titles, data_lst_segments)
        # else:
        #     print(f"File name {fichier} not found in the main excel.")
        # close_xlsm(app, wb, xlsm_recap_segments, the_file_exists, is_open, open_all)

    # def analyseAI():
        

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













excel_all_path = r".\data\benjamin_2_csv\Tableau récapitulatif - new algo"
excel_segments_path = r".\data\benjamin_2_csv\durée_segments"
recap_sheet_name = "Resultats_merged"
sheet_name_segment = "Feuil1"



# path = r"C:\Users\abarb\Documents\travail\stage et4\travail\codePlateau\data\A envoyer(pate a modeler)\A envoyer"
path = r"C:\Users\abarb\Documents\travail\stage et4\travail\codePlateau\data\A envoyer_antoine(non corrompue)\A envoyer"
date_folder = ""

path += "\\"
dossier = path + "xlsx"
dossier_graphique = path + "graph"
dossier_recap = path + r"recap\recap.xlsm"
dossier_recap_segments = path + r"recap\segments.xlsm"


# convert_csv_to_xlsx(path + "Expériences plateaux", dossier)

file_PlateauExp = "PlateauxExp.xlsx"
file_PlateauExp = "Plateaux_Exp.xlsx"
startCell_couverts = "F4"
startCell_couverts = "E4"

find_bites(dossier, dossier_graphique, date_folder, dossier_recap, dossier_recap_segments, file_PlateauExp, startCell_couverts, file = "180624_Dorian_Laura_P1.xlsx", writeFileNames = True)
# find_bites(dossier, dossier_graphique, date_folder, "14_05_Benjamin.xlsx")
