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

dossier = ("C:/Users/sebas/OneDrive/Desktop/ProgSmarTray/data")

fichiers = []
for f in os.listdir(dossier):
    if f.endswith(".xlsx"):
        fichiers.append(os.path.join(dossier, f))

dossier_graphique = "C:/Users/sebas/OneDrive/Desktop/ProgSmarTray/result"

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
    
    # Création des trames temporelles filtrées e
    fs = 1.0 / np.mean(np.diff(df["time"])) # Fréquence d'échantillonnage
    lowcut = 0.5 # Fréquence de coupure basse
    highcut = 1.0
    order = 4

    # # Fonction de filtrage band-stop
    # def butter_bandstop(lowcut, highcut, fs, order=4):
    #     nyq = 0.5 * fs
    #     low = lowcut / nyq
    #     high = highcut / nyq
    #     if(low<0 or high<0):
    #         print("Erreur: fréquence de coupure inférieure à 0")
    #     b, a = butter(order, [low, high], btype="bandstop")
    #     return b, a

    # b, a = butter_bandstop(lowcut, highcut, fs, order)
    # df["Ptot_filtered"] = filtfilt(b, a, df["Ptot"])
    # df = df.reset_index(drop=True)

    #Calcul du poids consommé
    flag = False
    debut = 0
    for i in range(10, len(df) - 1):
        if not flag and df["Ptot"].iloc[i] + 4 < df["Ptot"].shift(-1).iloc[i]:
            flag = True
            debut = df["Ptot"].iloc[i + 1]
    fin = df["Ptot"].iloc[-1]  # Prenez la dernière valeur de Ptot
    poids_consome = math.trunc(debut - fin)  # Utilisez math.trunc pour éliminer les décimales
    print(f"Le poids consommé pendant le repas est : {poids_consome}")

    #Calcul de la durée du repas 
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
    print(f"La durée du repas est : {temps_repas}")
        
    #calcul du temps d'activité et du nombre de bouchée
    pics = []
    inALoop = False
    x = 0
    activity_time = 0
    debut_pic = 0
    for i in range(10, len(df)-1):
        y = df["Ptot"].iloc[i]
        if not inALoop and y + 4 < df["Ptot"].shift(-1).iloc[i]:
            inALoop = True
            debut_pic = y
            pics.append([i])
            x = y - 4
        elif inALoop and x > y:
            activity_time += y - debut_pic
            pics[len(pics)-1].append(i)
            inALoop = False
            bouche += 1
    print(f"activity_time est : {activity_time}")
    ratio = activity_time / temps_repas
    print(f"Le ratio d'activité est : {ratio}")

    bouchee = np.append(bouchee, bouche)
    print(f"Le nombre de bouchée pendant le repas est : {bouche}")
    bouche = 0 

    # Création de graphiques avec Plotly

    fig = go.Figure()
    fig.update_layout(
        title_text="Poids Total en fonction du temps",
        xaxis_title="Temps en secondes",
        yaxis_title="Poids en grammes"
        )
    fig.add_trace(
        go.Scatter(y=df["Ptot"], x=df["time"], mode="lines", name="Poids Total")
    )
    
    # Enregistrement des graphiques
    filepath = os.path.join(
        dossier_graphique,
        "Graph_{}.html".format(os.path.basename(fichier).split(".")[0]),
    )
    fig.write_html(filepath)

