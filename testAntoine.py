import os
import pandas as pd
import numpy as np
from scipy.signal import butter, filtfilt
import plotly.graph_objects as go
from tkinter import filedialog, Tk

""""Variables globales"""
bouchee = np.array([])
bouche = 0
conso = np.array([])
tempsTot = np.array([])

# Sélection du dossier et lecture des fichiers

print("hello")
root = Tk()
root.withdraw()  # Pour ne pas afficher la fenêtre Tk


# Sélection du dossier contenant les données
# dossier = filedialog.askdirectory()
dossier = (
    r"C:\Users\comma\Documents\travail\Polytech\stage s8\gihtub\codePlateau\donneexslx\donneexslx"
)
fichiers = []
for f in os.listdir(dossier):
    if f.endswith(".xlsx"):
        fichiers.append(os.path.join(dossier, f))


# Sélection du dossier pour enregistrer les graphiques
# dossier_graphique = filedialog.askdirectory()
dossier_graphique = r"C:\Users\comma\Documents\travail\Polytech\stage s8\gihtub\codePlateau\donneexslx\donneexslx\diagramme"


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
    df.columns = ["time", "Ptot"]  # Assigner les noms de colonnes7
    df = df[df["Ptot"] > 100]  # Filtre sur le poids minimal de l'assiette
    df["Ptot"] = np.abs(df["Ptot"])
    df = df[df["Ptot"] < 3000] 
    df["time"] = df["time"] / 1000
    
    # Création des trames temporelles filtrées (à adapter en fonction des spécificités du signal)
    #fs = 1.0 / df["time"].diff().mean()  # Fréquence d'échantillonnage
    fs = 1.0 / np.mean(np.diff(df["time"]))
    lowcut = 0.5
    highcut = 1.0
    order = 4
    print(fs)


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

    # Calculer la condition sur tout le DataFrame
    """condition = df["Ptot"] < df["Ptot"].shift(-1)
    temp = df["Ptot"]

    # Incrémenter 'bouche' par le nombre de fois où la condition est vraie
    if(condition):
        if(temp > df["Ptot"]):
            bouche+=1
            
    bouchee = np.append(bouche, bouchee)
    bouche = 0"""

    df = df.reset_index(drop=True)

    y = 0
    
    pics = []
    inALoop = False
    x = 0
    for i in range(10,len(df) - 1):
        if not inALoop and df["Ptot"].iloc[i] + 4 < df["Ptot"].shift(-1).iloc[i]:
            inALoop = True
            pics.append([i, df["time"].iloc[i]])
            x = df["Ptot"].iloc[i] - 4
        elif inALoop and x > df["Ptot"].iloc[i]:
            y = df["Ptot"].iloc[i]
            pics[len(pics)-1] += [i, df["time"].iloc[i]]
            inALoop = False
            bouche += 1


    # for i in range(y, len(df) - 1):
    #     if df["Ptot"].iloc[i] < df["Ptot"].shift(-1).iloc[i]:
    #         #pics.append([i])
    #         x = df["Ptot"].iloc[i]
    #         found = False  # Indicateur si une valeur supérieure a été trouvée
    #         for j in range(i + 1, len(df) - 1):
    #             if x > df["Ptot"].iloc[j]:
    #                 y = df["Ptot"].iloc[j] 
    #                 found = True
    #                 #pics[len(pics)-1].append(j)
    #                 break
    #         if found:
    #             bouche += 1
    print("pics :", pics)

    # Ajoute le compteur à bouchee et réinitialise bouche
    bouchee = np.append(bouchee, bouche)
    bouche = 0

    # Création de graphiques avec Plotly

    fig = go.Figure()
    fig.add_trace(
        go.Scatter(y=df["Ptot"], x=df["time"], mode="lines", name="Poids Total")
    )
    

    # Enregistrement des graphiques
    filepath = os.path.join(
        dossier_graphique,
        "Graph_{}.html".format(os.path.basename(fichier).split(".")[0]),
    )
    fig.write_html(filepath)

    print(bouchee)
