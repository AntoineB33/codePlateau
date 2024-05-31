import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import plotly.graph_objs as go
from scipy.signal import butter, filtfilt
from openpyxl import load_workbook
from plotly.subplots import make_subplots

# Define the directory containing the data files
dossier = "c:/Users/comma/Documents/travail/Polytech/stage s8/gihtub/codePlateau/donneexslx/donneexslx"
dossier_graphique = "c:/Users/comma/Documents/travail/Polytech/stage s8/gihtub/codePlateau/R_code/result"

# List all files in the directory
fichiers = [os.path.join(dossier, f) for f in os.listdir(dossier) if f.endswith("5.xlsx")]

# Initialize the final dataframe
Tableau_Final = pd.DataFrame(columns=["Duree_Totale", "Poids_Conso", "Action", "Duree_activite_Totale",
                                      "Duree_activite_mean", "Duree_activite_max", "Duree_activite_min",
                                      "Proportion_activite_%", "Bouchees", "Num_fichier"])

# Functions
def butter_bandstop_filter(data, lowcut, highcut, fs, order):
    nyq = 0.5 * fs
    low = lowcut / nyq
    high = highcut / nyq
    b, a = butter(order, [low, high], btype='bandstop')
    y = filtfilt(b, a, data)
    return y

def find_inflexion_points(data):
    curvature = np.abs(np.diff(np.diff(data)))
    inflexion_indices = np.where(curvature > 5)[0]
    return inflexion_indices

def create_segments(indices, indice_min_duration):
    segments = []
    current_segment = [indices[0]]
    
    for i in range(1, len(indices)):
        if indices[i] - indices[i-1] <= indice_min_duration:
            current_segment.extend(range(indices[i-1] + 1, indices[i]))
        else:
            if len(current_segment) > indice_min_duration:
                segments.append(np.unique(np.sort(current_segment)).tolist())
            current_segment = [indices[i]]
    
    if len(current_segment) > indice_min_duration:
        segments.append(current_segment)
    
    return segments

def complete_segments(segments, max_diff):
    completed_segments = []
    
    for segment in segments:
        completed_segment = [segment[0]]
        for i in range(1, len(segment)):
            diff = segment[i] - segment[i-1]
            if diff <= max_diff:
                missing_indices = list(range(segment[i-1] + 1, segment[i]))
                completed_segment.extend(missing_indices)
            completed_segment.append(segment[i])
        
        completed_segments.append(sorted(set(completed_segment)))
    
    return completed_segments

def calculate_segment_durations_time(segments, time_data):
    durations = [time_data[segment[-1]] - time_data[segment[0]] if len(segment) > 1 else 0 for segment in segments]
    return durations

def calculate_segment_weight(segments, Ptot_data):
    Ptot_min = [min(Ptot_data[segment]) if len(segment) > 1 else 0 for segment in segments]
    return Ptot_min

def segment_consecutive(indices):
    segments = []
    current_segment = [indices[0]]
    
    for i in range(1, len(indices)):
        if indices[i] - indices[i-1] == 1:
            current_segment.append(indices[i])
        else:
            segments.append(current_segment)
            current_segment = [indices[i]]
    
    segments.append(current_segment)
    return segments


def dernier_indice_segment(segment_action, indice_bites):
    dernier_indices = []
    for indice_bite in indice_bites:
        if indice_bite < len(segment_action):
            dernier_indices.append(segment_action[indice_bite][-1])
        else:
            print(f"Warning: indice_bite {indice_bite} is out of range for segment_action of length {len(segment_action)}")
    return dernier_indices

# def dernier_indice_segment(segment_action, indice_bites):
#     dernier_indices = [segment_action[indice_bite][-1] for indice_bite in indice_bites]
#     return dernier_indices

# Process each file
for fichier in fichiers:
    # Read the data
    df = pd.read_excel(fichier)
    df.columns = ["time", "Ptot"]
    
    plate_weight_min = 100  # Minimum weight of the plate in grams
    data = df[df["Ptot"] > plate_weight_min]
    data["time"] = data["time"] / 1000
    
    if data.empty:
        continue
    
    min_bite_duration = 1  # Minimum bite duration in seconds
    min_bite_weight = 4    # Minimum bite weight in grams
    
    # Filter smoothing time
    int_time = 0.2
    indices_time = [data.index[0]]
    
    for i in range(1, len(data)):
        if data["time"].iloc[i] - data["time"].iloc[indices_time[-1]] >= int_time:
            indices_time.append(i)
    
    filtered_data = data.iloc[indices_time].reset_index(drop=True)
    
    if filtered_data.empty:
        continue
    
    # Filter smoothing weight
    seuil_poids = min_bite_weight
    filtered_data_wo_noise = filtered_data.copy()
    i = 0
    
    while i < len(filtered_data) - 1:
        val_ini = filtered_data_wo_noise["Ptot"].iloc[i]
        j = i + 1
        
        while j < len(filtered_data) and np.abs(filtered_data["Ptot"].iloc[j] - val_ini) < seuil_poids:
            j += 1
        
        if j < len(filtered_data):
            filtered_data_wo_noise.loc[i:j-1, "Ptot"] = val_ini
        i = j
    
    filtered_data = filtered_data_wo_noise
    
    indice_min_duration = np.argmax(np.cumsum(np.abs(np.diff(filtered_data["time"]))) >= min_bite_duration)
    
    ts_data_filt = butter_bandstop_filter(filtered_data["Ptot"], 0.5, 1, 1 / np.mean(np.diff(filtered_data["time"])), 4)
    inflexion_points = find_inflexion_points(ts_data_filt)
    segment_action = create_segments(inflexion_points, indice_min_duration)
    segment_action = complete_segments(segment_action, indice_min_duration)
    
    indice_action = np.unique(np.concatenate(segment_action))
    indice_non_action = np.setdiff1d(np.arange(len(filtered_data)), indice_action)
    segment_non_action = segment_consecutive(indice_non_action)
    
    durations_action_time = calculate_segment_durations_time(segment_action, filtered_data["time"])
    mean_action = np.mean(durations_action_time)
    max_action = np.max(durations_action_time)
    min_action = np.min(durations_action_time)
    duree_totale_action = np.sum(durations_action_time)
    
    duree_repas = np.abs(filtered_data["time"].iloc[segment_action[1][0]] - filtered_data["time"].iloc[segment_action[-1][-1]])
    nb_action = len(segment_action)
    proportion_action = duree_totale_action / duree_repas
    
    weight_non_action = calculate_segment_weight(segment_non_action, filtered_data["Ptot"])
    weight_non_action = [w for w in weight_non_action if w > plate_weight_min]
    poids_conso = weight_non_action[0] - min(weight_non_action)
    bites = len([w for w in np.diff(weight_non_action) if w < 0])
    indice_bites = [i + 1 for i in range(len(weight_non_action) - 1) if np.diff(weight_non_action)[i] < 0]
    
    temp_df = pd.DataFrame({
        "Duree_Totale": [duree_repas],
        "Poids_Conso": [poids_conso],
        "Action": [nb_action],
        "Duree_activite_Totale": [round(duree_totale_action, 3)],
        "Duree_activite_mean": [round(mean_action, 3)],
        "Duree_activite_max": [round(max_action, 3)],
        "Duree_activite_min": [round(min_action, 3)],
        "Proportion_activite_%": [round(proportion_action * 100, 3)],
        "Bouchees": [bites],
        "Num_fichier": [os.path.splitext(os.path.basename(fichier))[0]]
    })
    
    Tableau_Final = pd.concat([Tableau_Final, temp_df], ignore_index=True)
    
    indice_bites = dernier_indice_segment(segment_action, indice_bites)
    time_bites = filtered_data["time"].iloc[indice_bites]
    weight_bites = filtered_data["Ptot"].iloc[indice_bites]
    
    # Plotting
    fig = make_subplots()
    
    fig.add_trace(go.Scatter(x=filtered_data["time"], y=filtered_data["Ptot"], mode='lines', name='Données filtrées'))
    
    for i, segment in enumerate(segment_action):
        fig.add_trace(go.Scatter(x=filtered_data["time"].iloc[segment], y=filtered_data["Ptot"].iloc[segment],
                                 mode='lines', name=f'Action {i}', line=dict(color=f'rgb{plt.cm.rainbow(i / len(segment_action))[:3]}')))
    
    for i in range(len(time_bites)):
        fig.add_trace(go.Scatter(x=[time_bites.iloc[i], time_bites.iloc[i]], y=[filtered_data["Ptot"].min(), weight_bites.iloc[i]],
                                 mode='lines', name=f'Bouchée n°{i + 1}', line=dict(color='gray', dash='dot')))
    
    for i in range(len(time_bites)):
        fig.add_trace(go.Scatter(x=[time_bites.iloc[i], time_bites.iloc[i]], y=[filtered_data["Ptot"].min(), weight_bites.iloc[i]],
                                 mode='lines', name=f'Bouchée n°{i + 1}', line=dict(color='green', dash='dot')))


    # for i, segment in enumerate(segment_action):
    #     fig.add_trace(go.Scatter(x=filtered_data["time"].iloc[segment], y=filtered_data["Ptot"].iloc[segment],
    #                              mode='lines', name=f'Action {i}', line=dict(color=plt.cm.rainbow(i / len(segment_action)))))
    
    # for i in range(len(time_bites)):
    #     fig.add_trace(go.Scatter(x=[time_bites.iloc[i], time_bites.iloc[i]], y=[filtered_data["Ptot"].min(), weight_bites.iloc[i]],
    #                              mode='lines', name=f'Bouchée n°{i + 1}', line=dict(color='green', dash='dot')))
    
    # fig.add_trace(go.Scatter(x=filtered_data["time"], y=ts_data_filt, mode='lines', name='Analyse fréquentielle'))
    
    fig.update_layout(title=f'Repas : {os.path.splitext(os.path.basename(fichier))[0]}',
                      xaxis_title='Temps', yaxis_title='Ptot')
    
    fig.write_html(os.path.join(dossier_graphique, f'Graph_Repas_{os.path.splitext(os.path.basename(fichier))[0]}.html'))

    print(os.path.basename(fichier))

# Save the final table to a CSV file
Tableau_Final.to_csv(os.path.join(dossier_graphique, "Tableau_Final.csv"), index=False)
