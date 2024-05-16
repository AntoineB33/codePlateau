import os
import pandas as pd
import numpy as np
from scipy.signal import butter, filtfilt
import plotly.graph_objects as go
from tkinter import filedialog, Tk


# Sélection du dossier et lecture des fichiers

print("hello")
root = Tk()
root.withdraw()  # Pour ne pas afficher la fenêtre Tk


# Sélection du dossier contenant les données
# dossier = filedialog.askdirectory()
dossier = (
    "c:/Users/comma/Documents/travail/Polytech/stage s8/code/donneexslx/donneexslx"
)
fichiers = []
for f in os.listdir(dossier):
    if f.endswith(".xlsx"):
        fichiers.append(os.path.join(dossier, f))


# Sélection du dossier pour enregistrer les graphiques
# dossier_graphique = filedialog.askdirectory()
dossier_graphique = "c:/Users/comma/Documents/travail/Polytech/stage s8/code/donneexslx/donneexslx/diagramme"


# Création du DataFrame pour les résultats finaux


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
    df.columns = ["Ptot", "time"]  # Assigner les noms de colonnes
    df = df[df["Ptot"] > 100]  # Filtre sur le poids minimal de l'assiette

    # Création des trames temporelles filtrées (à adapter en fonction des spécificités du signal)
    fs = 1.0 / df["time"].diff().mean()  # Fréquence d'échantillonnage
    lowcut = 0.5
    highcut = 1.0
    order = 4


    # Fonction de filtrage band-stop
    def butter_bandstop(lowcut, highcut, fs, order=4):
        nyq = 0.5 * fs
        low = lowcut / nyq
        high = highcut / nyq
        if(low<0 or high<0):
            print("Erreur: fréquence de coupure inférieure à 0")
        b, a = butter(order, [low, high], btype="bandstop")
        return b, a

    b, a = butter_bandstop(lowcut, highcut, fs, order)
    df["Ptot_filtered"] = filtfilt(b, a, df["Ptot"])

    # Autres traitements et analyses...



    # Création de graphiques avec Plotly

    fig = go.Figure()
    fig.add_trace(
        go.Scatter(y=df["time"], x=df["Ptot"], mode="lines", name="Poids Total")
    )
    fig.add_trace(
        go.Scatter(
            y=df["time"], x=df["Ptot_filtered"], mode="lines", name="Poids Filtré"
        )
    )

    # Enregistrement des graphiques
    filepath = os.path.join(
        dossier_graphique,
        "Graph_{}.html".format(os.path.basename(fichier).split(".")[0]),
    )
    fig.write_html(filepath)
