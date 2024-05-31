dossier = r"C:\Users\comma\Documents\travail\Polytech\stage s8\gihtub\codePlateau\donneexslx\donneexslx"
dossier_graphique = r"C:\Users\comma\Documents\travail\Polytech\stage s8\gihtub\codePlateau\detechHorizontal"
""""/////////////////////////Import/////////////////////////"""
import os
import pandas as pd
import numpy as np
from scipy.signal import butter, filtfilt
import plotly.graph_objects as go
from tkinter import filedialog, Tk
import math

"""/////////////////////////////////////////////////////////"""


""""///////////////////Variables globales///////////////////"""
bouche = 0
poids_min = float("inf")
debut_time = None
fin_time = None
indice_debut = 0
indice_fin = 0
"""/////////////////////////////////////////////////////////"""

root = Tk()
root.withdraw()

# dossier = ("C:/Users/sebas/OneDrive/Desktop/ProgSmarTray/data")

fichiers = []
for f in os.listdir(dossier):
    if f.endswith("5.xlsx"):
        fichiers.append(os.path.join(dossier, f))

# dossier_graphique = "C:/Users/sebas/OneDrive/Desktop/ProgSmarTray/result"

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

# Traitement des fichiers

for fichier in fichiers:
    print(fichier)
    df = pd.read_excel(fichier)
    df.columns = ["time", "Ptot"]  # Assigner les noms de colonnes
    df = df[df["Ptot"] > 100]  # Filtre sur le poids minimal de l'assiette
    df["Ptot"] = np.abs(df["Ptot"])
    df = df[df["Ptot"] < 3000]  # Filtre sur le poids maximal de l'assiette
    df["time"] = df["time"] / 1000  # Conversion en secondes

    
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


import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.signal import find_peaks

# Load the data

# Plot the data for visualization
# plt.figure(figsize=(14, 6))
# plt.plot(df["time"], df["Ptot"])
# plt.xlabel("Time (en s)")
# plt.ylabel("Poid (en g)")
# plt.title("Poid vs Time")
# plt.show()

# Identify horizontal levels using a threshold on the derivative
threshold = 0.5  # You can adjust this value based on your data
derivative = np.abs(np.diff(df["Ptot"]))
horizontal_indices = np.where(derivative < threshold)[0]

# Group the horizontal indices into levels
horizontal_levels = []
current_level = [horizontal_indices[0]]
for i in range(1, len(horizontal_indices)):
    if horizontal_indices[i] - horizontal_indices[i - 1] == 1:
        current_level.append(horizontal_indices[i])
    else:
        horizontal_levels.append(current_level)
        current_level = [horizontal_indices[i]]
horizontal_levels.append(current_level)

# Find the mean value for each level
levels = [np.mean(df["Ptot"].iloc[level]) for level in horizontal_levels]

# Detect spikes
peaks, _ = find_peaks(
    df["Ptot"], distance=5
)  # Adjust distance parameter based on your data
spikes = df.iloc[peaks]

# Find time intervals of spikes between two successive levels
time_intervals = []
for i in range(1, len(horizontal_levels)):
    start_time = df["time"].iloc[horizontal_levels[i - 1][-1]]
    end_time = df["time"].iloc[horizontal_levels[i][0]]
    time_intervals.append((start_time, end_time))

# Output the results
print("Horizontal Levels (in Poid):", levels)
print("Time Intervals of Spikes (in seconds):", time_intervals)

# Plot the results
# plt.figure(figsize=(14, 6))
# plt.plot(df["time"], df["Ptot"], label="Poid vs Time")
# for level in horizontal_levels:
#     plt.hlines(
#         df["Ptot"].iloc[level[0]],
#         df["time"].iloc[level[0]],
#         df["time"].iloc[level[-1]],
#         colors="r",
#         linestyles="dashed",
#     )
# for start_time, end_time in time_intervals:
#     plt.axvspan(start_time, end_time, color="yellow", alpha=0.3)
# plt.scatter(spikes["time"], spikes["Ptot"], color="g", label="Spikes")
# plt.xlabel("Time (en s)")
# plt.ylabel("Poid (en g)")
# plt.title("Poid vs Time with Detected Levels and Spikes")
# plt.legend()
# plt.show()








fig = go.Figure()
fig.update_layout(
    title_text="Poids Total en fonction du temps",
    xaxis_title="Temps (en s)",
    yaxis_title="Poids (en g)"
    )
fig.add_trace(
    go.Scatter(y=df["Ptot"], x=df["time"], mode="lines", name="Poids Total")
)

# Add vertical lines
max_value = df["Ptot"].max()
for i in range(len(horizontal_levels)-1):
    if df["Ptot"].iloc[horizontal_levels[i][-1]] > df["Ptot"].iloc[horizontal_levels[i+1][0]]:
        # Add a shape (rectangle) to fill the area between the vertical lines
        fig.add_shape(
            type="rect",
            x0=df["time"].iloc[horizontal_levels[i][-1]],
            x1=df["time"].iloc[horizontal_levels[i+1][0]],
            y0=0,  # Adjust y0 and y1 as needed
            y1=max_value,  # Adjust y0 and y1 as needed
            fillcolor="rgba(255, 0, 0, 0.2)",  # Transparent red fill color
            line=dict(color="rgba(255, 0, 0, 0)")  # No border line
        )


# Enregistrement des graphiques
filepath = os.path.join(
    dossier_graphique,
    "Graph_{}.html".format(os.path.basename(fichier).split(".")[0]),
)

fig.add_annotation(
text=f"Nombre de bouchées: {bouche}",
xref="paper",  # Utilise les coordonnées relatives à la zone du graphique
yref="paper",
x=0.98,  # Position sur l'axe des x, 1 étant tout à droite
y=0.98,  # Position sur l'axe des y, 1 étant tout en haut
showarrow=False,
font=dict(
    family="Calibri, monospace",
    size=16,
    color="#000000"
))

fig.write_html(filepath)
