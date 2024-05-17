dossier = r"C:\Users\comma\Documents\travail\Polytech\stage s8\gihtub\codePlateau\donneexslx\donneexslx"
dossier_graphique = r"C:\Users\comma\Documents\travail\Polytech\stage s8\gihtub\codePlateau\donneexslx\donneexslx\diagramme"
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


import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.signal import find_peaks

# Load the data

# Plot the data for visualization
plt.figure(figsize=(14, 6))
plt.plot(df["time"], df["Ptot"])
plt.xlabel("Time (en s)")
plt.ylabel("Poid (en g)")
plt.title("Poid vs Time")
plt.show()

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
plt.figure(figsize=(14, 6))
plt.plot(df["time"], df["Ptot"], label="Poid vs Time")
for level in horizontal_levels:
    plt.hlines(
        df["Ptot"].iloc[level[0]],
        df["time"].iloc[level[0]],
        df["time"].iloc[level[-1]],
        colors="r",
        linestyles="dashed",
    )
for start_time, end_time in time_intervals:
    plt.axvspan(start_time, end_time, color="yellow", alpha=0.3)
plt.scatter(spikes["time"], spikes["Ptot"], color="g", label="Spikes")
plt.xlabel("Time (en s)")
plt.ylabel("Poid (en g)")
plt.title("Poid vs Time with Detected Levels and Spikes")
plt.legend()
plt.show()
