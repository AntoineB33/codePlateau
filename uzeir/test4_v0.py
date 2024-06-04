import os
import pandas as pd
import numpy as np
import math
from plotly.graph_objects import Figure, Scatter
from tkinter import filedialog, Tk
from scipy.signal import find_peaks, peak_prominences

""""///////////////////Variables globales///////////////////"""
bouchees = 0
poids_min = float('inf')
debut_time = None
fin_time = None
indice_debut = 0
indice_fin = 0
"""/////////////////////////////////////////////////////////"""

root = Tk()
root.withdraw()

dossier = "./filtered_data"

fichiers = []
for f in os.listdir(dossier):
    if f.endswith(".xlsx") and f == "17.xlsx":
        fichiers.append(os.path.join(dossier, f))

dossier_graphique = "./uzeir/result"

Tableau_Final = pd.DataFrame(
    columns=[
        "Duree_Totale",
        "Poids_Conso",
        "Action",
        "Duree_activite_Totale",
        "Duree_activite_mean",
        "Duree_activite_max",
        "Duree_activite_min",
        "Proportion_activite_%",
        "Bouchees",
        "Num_fichier",
    ]
)

print(dossier)

# Fonction pour convertir le temps en minutes et secondes
def convert_time(seconds):
    minutes = seconds // 60
    seconds = seconds % 60
    return f"{minutes} min {seconds} s"

# Traitement des fichiers
for fichier in fichiers:
    print(fichier)
    df = pd.read_excel(fichier)
    df.columns = ["time", "Ptot"]
    df = df[df["Ptot"] > 100]
    df["Ptot"] = np.abs(df["Ptot"])
    df = df[df["Ptot"] < 3000]
    df["time"] = df["time"] / 1000




    # Filter smoothing time
    int_time = 0.2  # Time interval threshold
    indices_time = [df.index[0]]

    for i in range(1, len(df)):
        if df["time"].iloc[i] - df["time"].iloc[indices_time[-1]] >= int_time:
            indices_time.append(i)

    df = df.iloc[indices_time].reset_index(drop=True)

    if df.empty:
        continue  # Skip further processing if no data remains after filtering
    
    # Filter smoothing weight
    seuil_poids = 10  # Assuming a threshold, adjust as necessary
    filtered_df = df.copy()
    i = 0

    while i < len(df) - 1:
        val_ini = filtered_df["Ptot"].iloc[i]
        j = i + 1

        while j < len(df) and np.abs(df["Ptot"].iloc[j] - val_ini) < seuil_poids:
            j += 1

        if j < len(df):
            filtered_df.loc[i:j-1, "Ptot"] = val_ini
        i = j

    df = filtered_df  # Update df with the noise-filtered data




    # Calcul du poids consommé
    flag = False
    debut = 0
    for i in range(10, len(df) - 1):
        if not flag and df["Ptot"].iloc[i] + 4 < df["Ptot"].shift(-1).iloc[i]:
            flag = True
            debut = df["Ptot"].iloc[i + 1]
    fin = df["Ptot"].iloc[-1]
    poids_consome = math.trunc(debut - fin)
    print(f"Le poids consommé pendant le repas est : {poids_consome}")

    # Calcul de la durée du repas
    for i in range(len(df) - 1):
        if df["Ptot"].iloc[i] > 700 and (df["Ptot"].iloc[i] + 4) < df["Ptot"].shift(-1).iloc[i]:
            debut_time = df["time"].iloc[i + 1]
            indice_debut = i
            for j in range(i + 1, len(df)):
                if df["Ptot"].iloc[j] < poids_min:
                    poids_min = df["Ptot"].iloc[j]
                    fin_time = df['time'].iloc[j]
                    indice_fin = i
            break
    temps_repas = math.trunc(fin_time - debut_time)
    print(f"La durée du repas est : {convert_time(temps_repas)}")

    # Calcul du temps d'activité et du nombre de bouchée
    activity_time = 0
    peaks_x = []
    peaks_y = []

    # Peak detection
    peaks, _ = find_peaks(df["Ptot"], height=100, distance=5)  # Reduced distance for more sensitivity
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
        if (idx - 1 >= 0 and idx + 1 < len(df)) and \
           (df["Ptot"].iloc[idx] - df["Ptot"].iloc[idx - 1] > weight_diff_threshold) and \
           (df["Ptot"].iloc[idx] - df["Ptot"].iloc[idx + 1] > weight_diff_threshold):
            # Check if weight changes significantly before and after the peak
            window = 5  # Define a window around the peak to check weight change
            if (idx - window >= 0 and idx + window < len(df)) and \
               (abs(df["Ptot"].iloc[idx - window] - df["Ptot"].iloc[idx]) > weight_diff_threshold) and \
               (abs(df["Ptot"].iloc[idx + window] - df["Ptot"].iloc[idx]) > weight_diff_threshold):
                valid_peaks_indices.append(idx)

    valid_peaks = np.array(valid_peaks_indices)
    valid_prominences = peak_prominences(df["Ptot"], valid_peaks)[0]

    # Filter closely spaced peaks and merge windows
    min_diff = 50  # Minimum difference in indices between consecutive peaks
    final_peaks_indices = []
    merged_windows = []
    activity_time = 0
    
    for i in range(len(valid_peaks)):
        if not final_peaks_indices:
            final_peaks_indices.append(valid_peaks[i])
            window_start = valid_peaks[i]
            window_end = valid_peaks[i]
        else:
            last_peak_idx = final_peaks_indices[-1]
            if (valid_peaks[i] - last_peak_idx) > min_diff:
                final_peaks_indices.append(valid_peaks[i])
                merged_windows.append((window_start, window_end))
                activity_time += df["time"].iloc[window_end] - df["time"].iloc[window_start]
                window_start = valid_peaks[i]
                window_end = valid_peaks[i]
            else:
                # Merge peaks if they are close
                window_end = valid_peaks[i]
                if valid_prominences[i] > valid_prominences[np.where(valid_peaks == last_peak_idx)[0][0]]:
                    final_peaks_indices[-1] = valid_peaks[i]
        # Ensure the window includes the whole peak
        window_start = min(window_start, valid_peaks[i])
        window_end = max(window_end, valid_peaks[i])
    
    # Add the last window
    merged_windows.append((window_start, window_end))

    significant_peaks_x = df["time"].iloc[final_peaks_indices].values
    significant_peaks_y = df["Ptot"].iloc[final_peaks_indices].values

    bouchees = len(significant_peaks_y)  # Nombre de bouchées est le nombre de pics significatifs

    # Calcul du temps d'activité
    # if len(significant_peaks_x) > 0:
    #     activity_start_time = significant_peaks_x[0]
    #     activity_end_time = significant_peaks_x[-1]
    #     activity_time = math.trunc(activity_end_time - activity_start_time)
    # else:
    #     activity_time = 0

    ratio = activity_time / temps_repas if temps_repas > 0 else 0
    print(f"Le ratio d'activité est : {ratio}")
    print(f"Le nombre de bouchée pendant le repas est : {bouchees}")
    print(f"Le temps d'activité est : {convert_time(activity_time)}")

    # Création de graphiques avec Plotly
    fig = Figure()
    fig.update_layout(
        title_text="Poids Total en fonction du temps",
        xaxis_title="Temps en secondes",
        yaxis_title="Poids en grammes"
    )
    fig.add_trace(
        Scatter(y=df["Ptot"], x=df["time"], mode="lines", name="Poids Total")
    )

    # Ajouter les fenêtres en rouge pour chaque pic
    for start_idx, end_idx in merged_windows:
        # Début de la fenêtre : recherche en arrière jusqu'à ce que le poids commence à augmenter
        while start_idx > 0 and df["Ptot"].iloc[start_idx - 1] < df["Ptot"].iloc[start_idx]:
            start_idx -= 1
        
        # Fin de la fenêtre : recherche en avant jusqu'à ce que le poids redevienne constant
        while end_idx < len(df) - 1 and df["Ptot"].iloc[end_idx + 1] < df["Ptot"].iloc[end_idx]:
            end_idx += 1

        fig.add_shape(
            type="rect",
            x0=df["time"].iloc[start_idx], y0=df["Ptot"].min(),
            x1=df["time"].iloc[end_idx], y1=df["Ptot"].max(),
            line=dict(color="Red", width=2),
            fillcolor="Red",
            opacity=0.2,
        )

    fig.add_trace(
        Scatter(y=significant_peaks_y, x=significant_peaks_x, mode="markers", name="Pics détectés", marker=dict(color='red', size=8))
    )
    fig.add_annotation(
        x=df["time"].iloc[-1], y=df["Ptot"].max(),
        text=f"Poids consommé: {poids_consome} g<br>Nombre de bouchées: {bouchees}<br>Temps d'activité: {convert_time(activity_time)}<br>Ratio d'activité: {ratio:.2f}",
        showarrow=False,
        align='left',
        xanchor='right',
        yanchor='top'
    )

    # Enregistrement des graphiques
    filepath = os.path.join(
        dossier_graphique,
        "Graph_{}_liss.html".format(os.path.basename(fichier).split(".")[0]),
    )
    fig.write_html(filepath)