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
poids_min = float('inf')
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
    df = df[df["Ptot"] < 3000] # Filtre sur le poids maximal de l'assiette
    df["time"] = df["time"] / 1000 # Conversion en secondes
    
import pandas as pd
import ruptures as rpt
import matplotlib.pyplot as plt

# Charger les données
# df = pd.read_csv('path_to_your_data.csv')
# Comme nous n'avons pas le fichier CSV exact, supposons que les données sont déjà dans df
# df.columns = ["time", "Ptot"]

# Convertir les données en format requis par ruptures (numpy array)
signal = df['Ptot'].values

# Créer un modèle de changement
model = rpt.Pelt(model="rbf").fit(signal)

# Détecter les points de changement
breakpoints = model.predict(pen=10)

# Afficher les résultats
print("Indices des points de changement :", breakpoints)

# Tracer les résultats
rpt.display(signal, breakpoints)
plt.xlabel('Temps (en s)')
plt.ylabel('Poids (en g)')
plt.title('Détection des paliers')
plt.show()
