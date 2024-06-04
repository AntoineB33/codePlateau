import os
import pandas as pd
import numpy as np
from scipy.signal import butter, filtfilt
import plotly.graph_objects as go
import plotly.io as pio

# Define directories
data_folder = "c:/Users/comma/Documents/travail/Polytech/stage s8/gihtub/codePlateau/donneexslx/donneexslx"
output_folder = "c:/Users/comma/Documents/travail/Polytech/stage s8/gihtub/codePlateau/R_code/result_R"

# Function to apply bandstop filter
def butter_bandstop_filter(data, lowcut, highcut, fs, order):
    nyq = 0.5 * fs
    low = lowcut / nyq
    high = highcut / nyq
    b, a = butter(order, [low, high], btype='bandstop')
    y = filtfilt(b, a, data)
    return y

# Function to find inflexion points
def find_inflexion_points(data):
    curvature = np.abs(np.diff(np.diff(data)))
    inflexion_indices = np.where(curvature > 5)[0]  # Adjust the threshold as needed
    return inflexion_indices

# Function to create segments
def create_segments(indices, min_duration):
    segments = []
    current_segment = [indices[0]]
    for i in range(1, len(indices)):
        if indices[i] - indices[i - 1] <= min_duration:
            current_segment.extend(range(indices[i - 1] + 1, indices[i]))
        else:
            if len(current_segment) > min_duration:
                segments.append(sorted(set(current_segment)))
            current_segment = [indices[i]]
    if len(current_segment) > min_duration:
        segments.append(current_segment)
    return segments

# Function to complete segments
def complete_segments(segments, max_diff):
    completed_segments = []
    for segment in segments:
        completed_segment = [segment[0]]
        for i in range(1, len(segment)):
            diff = segment[i] - segment[i - 1]
            if diff <= max_diff:
                missing_indices = list(range(segment[i - 1] + 1, segment[i]))
                completed_segment.extend(missing_indices)
            completed_segment.append(segment[i])
        completed_segments.append(sorted(set(completed_segment)))
    return completed_segments

# Function to calculate segment durations
def calculate_segment_durations(segments, time_data):
    durations = [time_data[seg[-1]] - time_data[seg[0]] if len(seg) > 1 else 0 for seg in segments]
    return durations

# Function to calculate segment weights
def calculate_segment_weight(segments, Ptot_data):
    weights = [min(Ptot_data[seg]) if len(seg) > 1 else 0 for seg in segments]
    return weights

# Function to segment consecutive indices
def segment_consecutive(indices):
    segments = []
    current_segment = [indices[0]]
    for i in range(1, len(indices)):
        if indices[i] - indices[i - 1] == 1:
            current_segment.append(indices[i])
        else:
            segments.append(current_segment)
            current_segment = [indices[i]]
    segments.append(current_segment)
    return segments

# Function to get the last index in each segment
def last_index_in_segments(segment_action, bite_indices):
    last_indices = [segment_action[idx][-1] for idx in bite_indices]
    return last_indices

# Initialize final table
Tableau_Final = pd.DataFrame(columns=["Duree_Totale", "Poids_Conso", "Action", "Duree_activite_Totale", 
                                      "Duree_activite_mean", "Duree_activite_max", "Duree_activite_min", 
                                      "Proportion_activite_%", "Bouchees", "Num_fichier"])

# Get list of files
files = [os.path.join(data_folder, f) for f in os.listdir(data_folder) if f == '5.xlsx']

# Process each file
for file in files:
    df = pd.read_excel(file)
    df.columns = ["time", "Ptot"]

    plate_weight_min = 100  # grams
    data = df[df['Ptot'] > plate_weight_min]
    data['time'] /= 1000  # Convert to seconds

    if data.empty:
        continue

    min_bite_duration = 1  # seconds
    min_bite_weight = 4  # grams

    int_time = 0.2  # seconds
    indices_time = [data.index[0]]

    for i in range(1, len(data)):
        if data['time'].iloc[i] - data['time'].iloc[indices_time[-1]] >= int_time:
            indices_time.append(i)

    filtered_data = data.iloc[indices_time]

    row_indices = filtered_data.index
    filtered_data_wo_noise = filtered_data.copy()
    for i in range(1, len(filtered_data)):
        if abs(filtered_data.iloc[i]['Ptot'] - filtered_data.iloc[i - 1]['Ptot']) < min_bite_weight:
            filtered_data_wo_noise.at[row_indices[i], 'Ptot'] = filtered_data.iloc[i - 1]['Ptot']

    fs = 1 / np.mean(np.diff(filtered_data['time']))
    ts_data_filt = butter_bandstop_filter(filtered_data_wo_noise['Ptot'], 0.5, 1, fs, 4)

    inflexion_points = find_inflexion_points(ts_data_filt)
    segment_action = create_segments(inflexion_points, int_time / 0.2)
    segment_action = complete_segments(segment_action, int_time / 0.2)

    action_indices = np.hstack(segment_action)
    non_action_indices = np.setdiff1d(filtered_data.index, action_indices)
    segment_non_action = segment_consecutive(non_action_indices)

    durations_action_time = calculate_segment_durations(segment_action, filtered_data['time'])
    mean_action = np.mean(durations_action_time)
    max_action = np.max(durations_action_time)
    min_action = np.min(durations_action_time)
    total_action_duration = np.sum(durations_action_time)
    meal_duration = filtered_data['time'].iloc[segment_action[-1][-1]] - filtered_data['time'].iloc[segment_action[1][0]]
    num_actions = len(segment_action)
    proportion_action = total_action_duration / meal_duration

    weights_non_action = calculate_segment_weight(segment_non_action, filtered_data['Ptot'])
    weights_non_action = weights_non_action[weights_non_action > plate_weight_min]
    consumed_weight = weights_non_action[0] - min(weights_non_action)
    bites = np.sum(np.diff(weights_non_action) < 0)
    bite_indices = np.where(np.diff(weights_non_action) < 0)[0] + 1

    temp_df = pd.DataFrame({
        "Duree_Totale": [meal_duration],
        "Poids_Conso": [consumed_weight],
        "Action": [num_actions],
        "Duree_activite_Totale": [round(total_action_duration, 3)],
        "Duree_activite_mean": [round(mean_action, 3)],
        "Duree_activite_max": [round(max_action, 3)],
        "Duree_activite_min": [round(min_action, 3)],
        "Proportion_activite_%": [round(proportion_action * 100, 3)],
        "Bouchees": [bites],
        "Num_fichier": [os.path.basename(file).replace('.xlsx', '')]
    })

    Tableau_Final = pd.concat([Tableau_Final, temp_df], ignore_index=True)

    bite_indices = last_index_in_segments(segment_action, bite_indices)
    time_bites = filtered_data['time'].iloc[bite_indices]
    weight_bites = filtered_data['Ptot'].iloc[bite_indices]

    colors = [f'rgba({np.random.randint(0,255)},{np.random.randint(0,255)},{np.random.randint(0,255)},1)' for _ in segment_action]

    fig = go.Figure()
    fig.add_trace(go.Scatter(x=filtered_data['time'], y=filtered_data['Ptot'], mode='lines', name='Données filtrées', line=dict(color='black')))

    for i, segment in enumerate(segment_action):
        segment_data = filtered_data.iloc[segment]
        fig.add_trace(go.Scatter(x=segment_data['time'], y=segment_data['Ptot'], mode='lines', name=f'Action {i+1}', line=dict(color=colors[i % len(colors)])))

    for i in range(len(time_bites)):
        fig.add_trace(go.Scatter(x=[time_bites.iloc[i], time_bites.iloc[i]], y=[filtered_data['Ptot'].min(), weight_bites.iloc[i]],
                                 mode='lines', line=dict(color='green', dash='dot'), name=f'Bouchée n°{i+1}'))

    fig.add_trace(go.Scatter(x=filtered_data['time'], y=ts_data_filt, mode='lines', name='Analyse fréquentielle', line=dict(color='black')))
    fig.update_layout(title=f"Repas : {os.path.basename(file).replace('.xlsx', '')}",
                      xaxis_title='Temps', yaxis_title='Ptot')

    pio.write_html(fig, file=os.path.join(output_folder, f"Graph_Repas_{os.path.basename(file).replace('.xlsx', '')}.html"))

    print(os.path.basename(file))

print(Tableau_Final)
