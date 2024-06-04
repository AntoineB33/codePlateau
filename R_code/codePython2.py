import os
import pandas as pd
import numpy as np
from scipy.signal import butter, filtfilt
import plotly.graph_objs as go
from plotly.subplots import make_subplots
import openpyxl

dossier = "c:/Users/comma/Documents/travail/Polytech/stage s8/gihtub/codePlateau/donneexslx/donneexslx"
fichiers = [os.path.join(dossier, f) for f in os.listdir(dossier) if f == "7.xlsx"]
dossier_graphique = "c:/Users/comma/Documents/travail/Polytech/stage s8/gihtub/codePlateau/R_code/result"

Tableau_Final = pd.DataFrame(columns=[
    'Duree_Totale', 'Poids_Conso', 'Action', 'Duree_activite_Totale', 
    'Duree_activite_mean', 'Duree_activite_max', 'Duree_activite_min', 
    'Proportion_activite_%', 'Bouchees', 'Num_fichier'
])

# Function definitions
def butter_bandstop_filter(data, lowcut, highcut, fs, order):
    nyq = 0.5 * fs
    low = lowcut / nyq
    high = highcut / nyq
    
    b, a = butter(order, [low, high], btype='bandstop')
    y = filtfilt(b, a, data)
    
    return y

def find_inflexion_points(data):
    curvature = np.abs(np.diff(np.diff(data)))
    inflexion_indices = np.where(curvature > 5)[0]  # Adjust threshold as needed
    return inflexion_indices

def create_segments(indices, min_duration):
    segments = []
    current_segment = [indices[0]]
    
    for i in range(1, len(indices)):
        if indices[i] - indices[i-1] <= min_duration:
            current_segment.extend(range(indices[i-1] + 1, indices[i]))
        else:
            if len(current_segment) > min_duration:
                segments.append(sorted(set(current_segment)))
            current_segment = [indices[i]]
    
    if len(current_segment) > min_duration:
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
    durations = [time_data[seg[-1]] - time_data[seg[0]] if len(seg) > 1 else 0 for seg in segments]
    return durations

def calculate_segment_weight(segments, Ptot_data):
    Ptot_min = [min(Ptot_data[seg]) if len(seg) > 1 else 0 for seg in segments]
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
    return [segment_action[i][-1] for i in indice_bites]

# Processing files
for fichier in fichiers:
    df = pd.read_excel(fichier, header=None)
    df.columns = ["time", "Ptot"]
    
    plate_weight_min = 100  # in grams
    data = df[df['Ptot'] > plate_weight_min]
    data['time'] = data['time'] / 1000
    
    if data.empty:
        continue
    
    min_bite_duration = 1  # in seconds
    min_bite_weight = 4  # in grams
    
    int_time = 0.2  # minimum time interval
    indices_time = [0]
    
    while True:
        Cumul_duration = np.cumsum(np.diff(data['time'].values, prepend=0))
        indice = np.where(Cumul_duration >= int_time)[0]
        if len(indice) == 0:
            break
        indices_time.append(indice[0])
        data = data.iloc[indice[0]:]
    
    indices_time = list(set(indices_time))
    filtered_data = df.iloc[indices_time]
    
    seuil_poids = min_bite_weight
    filtered_data_wo_noise = filtered_data.copy()
    
    while True:
        val_ini = filtered_data.iloc[0]['Ptot']
        indice = np.where(np.abs(filtered_data['Ptot'] - val_ini) >= seuil_poids)[0]
        if len(indice) == 0:
            break
        indice = indice[0]
        filtered_data_wo_noise['Ptot'].iloc[:indice] = val_ini
        filtered_data = filtered_data.iloc[indice:]
    
    filtered_data = filtered_data_wo_noise
    
    indice_min_duration = np.where(np.cumsum(np.diff(filtered_data['time'].values, prepend=0)) >= min_bite_duration)[0][0]
    
    ts_data_filt = butter_bandstop_filter(filtered_data['Ptot'].values, 0.5, 1, 1 / np.mean(np.diff(filtered_data['time'].values)), 4)
    
    inflexion_points = find_inflexion_points(ts_data_filt)
    segment_action = create_segments(inflexion_points, indice_min_duration)
    segment_action = complete_segments(segment_action, indice_min_duration)
    
    indice_action = np.concatenate(segment_action)
    indice_non_action = np.setdiff1d(range(len(filtered_data)), indice_action)
    segment_non_action = segment_consecutive(indice_non_action)
    
    durations_action_time = calculate_segment_durations_time(segment_action, filtered_data['time'].values)
    mean_action = np.mean(durations_action_time)
    max_action = np.max(durations_action_time)
    min_action = np.min(durations_action_time)
    duree_totale_action = np.sum(durations_action_time)
    duree_repas = filtered_data['time'].iloc[segment_action[1][0]] - filtered_data['time'].iloc[segment_action[-1][-1]]
    nb_action = len(segment_action)
    proportion_action = duree_totale_action / duree_repas
    
    weight_non_action = calculate_segment_weight(segment_non_action, filtered_data['Ptot'].values)
    weight_non_action = [w for w in weight_non_action if w > plate_weight_min]
    poids_conso = weight_non_action[0] - np.min(weight_non_action)
    bites = len([w for w in np.diff(weight_non_action) if w < 0])
    indice_bites = [i + 1 for i, w in enumerate(np.diff(weight_non_action)) if w < 0]
    
    temp_df = pd.DataFrame([{
        'Duree_Totale': duree_repas,
        'Poids_Conso': poids_conso,
        'Action': nb_action,
        'Duree_activite_Totale': round(duree_totale_action, 3),
        'Duree_activite_mean': round(mean_action, 3),
        'Duree_activite_max': round(max_action, 3),
        'Duree_activite_min': round(min_action, 3),
        'Proportion_activite_%': round(proportion_action * 100, 3),
        'Bouchees': bites,
        'Num_fichier': os.path.splitext(os.path.basename(fichier))[0]
    }])
    
    Tableau_Final = pd.concat([Tableau_Final, temp_df], ignore_index=True)
    
    indice_bites = dernier_indice_segment(segment_action, indice_bites)
    time_bites = filtered_data['time'].values[indice_bites]
    weight_bites = filtered_data['Ptot'].values[indice_bites]
    
    fig = make_subplots(rows=1, cols=1)
    
    fig.add_trace(go.Scatter(x=filtered_data['time'], y=filtered_data['Ptot'], mode='lines', name='Données filtrées', line=dict(color='black')))
    
    colors = plt.cm.rainbow(np.linspace(0, 1, len(segment_action)))
    
    for i, segment in enumerate(segment_action):
        segment_data = filtered_data.iloc[segment]
        fig.add_trace(go.Scatter(x=segment_data['time'], y=segment_data['Ptot'], mode='lines', name=f'Action {i+1}', line=dict(color=f'rgba({colors[i][0]}, {colors[i][1]}, {colors[i][2]}, 1)')))
    
    for i, tb in enumerate(time_bites):
        fig.add_trace(go.Scatter(x=[tb, tb], y=[filtered_data['Ptot'].min(), weight_bites[i]], mode='lines', name=f'Bouchée n° {i+1}', line=dict(color='green', dash='dot')))
    
    fig.add_trace(go.Scatter(x=filtered_data['time'], y=ts_data_filt, mode='lines', name='Analyse fréquentielle', line=dict(color='black')))
    
    fig.update_layout(title=f'Repas : {os.path.splitext(os.path.basename(fichier))[0]}', xaxis_title='Temps', yaxis_title='Ptot')
    
    fig.write_html(os.path.join(dossier_graphique, f"Graph_Repas_{os.path.splitext(os.path.basename(fichier))[0]}.html"))
    
    print(os.path.basename(fichier))

print(Tableau_Final)
