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
min_bite_weight = 4
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


# Fonction de filtrage passe-bas
def butter_lowpass(cutoff, fs, order=4):
    nyq = 0.5 * fs
    normal_cutoff = cutoff / nyq
    b, a = butter(order, normal_cutoff, btype='low', analog=False)
    return b, a

def lowpass_filter(data, cutoff, fs, order=4):
    b, a = butter_lowpass(cutoff, fs, order=order)
    y = filtfilt(b, a, data)
    return y

# Traitement des fichiers

for fichier in fichiers:
    print(fichier)
    df = pd.read_excel(fichier)
    df.columns = ["time", "Ptot"]  # Assigner les noms de colonnes
    df = df[df["Ptot"] > 100]  # Filtre sur le poids minimal de l'assiette
    df["Ptot"] = np.abs(df["Ptot"]) 
    df = df[df["Ptot"] < 3000] # Filtre sur le poids maximal de l'assiette
    df["time"] = df["time"] / 1000 # Conversion en secondes
    
    
    # Application du filtre passe-bas
    fs = 1.0 / np.mean(np.diff(df["time"]))  # Fréquence d'échantillonnage
    # cutoff = 0.1  # Fréquence de coupure pour le filtre passe-bas
    # df["Ptot"] = lowpass_filter(df["Ptot"], cutoff, fs)
    lowcut = 0.5 # Fréquence de coupure basse
    highcut = 1.0
    order = 4


    seuil_poids = min_bite_weight

    filtered_data['Index'] = np.arange(len(filtered_data))
    filtered_data_true = filtered_data.copy()
    filtered_data_wo_noise = filtered_data.copy()
    n_true = np.arange(len(filtered_data_true))

    indice = 0

    while not pd.isna(indice):
        val_ini = float(filtered_data_true.iloc[0]['Ptot'])

        indices = np.where(np.abs(filtered_data_true['Ptot'] - val_ini) >= seuil_poids)[0]
        if len(indices) > 0:
            indice = indices[0] - 1
        else:
            indice = np.nan
        
        if not pd.isna(indice):
            n_true = n_true[np.where(n_true == indice)[0][0]:]
            indice_commun = np.where(filtered_data_true.iloc[0]['time'] == filtered_data_wo_noise['time'])[0][0]

            filtered_data_wo_noise.loc[indice_commun:indice, 'Ptot'] = val_ini
            n_true = n_true[1:]  
            filtered_data_true = filtered_data_true[filtered_data_true['Index'] == n_true[0]:]

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
    indice_debut = 0
    indice_fin = len(df) - 1
    for i in range(len(df) - 1):
        if df["Ptot"].iloc[i] > 700 and (df["Ptot"].iloc[i] + 4) < df["Ptot"].shift(-1).iloc[i]:
            debut_time = df["time"].iloc[i + 1]
            indice_debut = i
            for j in range(i + 1, len(df)):
                if df["Ptot"].iloc[j] < poids_min:
                    poids_min = df["Ptot"].iloc[j]
                    fin_time = df['time'].iloc[j]
                    indice_fin = j
            break
    temps_repas = math.trunc(fin_time - debut_time)
    tempsmin = (temps_repas // 60)
    tempssec = (temps_repas % 60)

    print(f"La durée du repas est : {tempsmin} minutes et {tempssec} secondes")
        
    #calcul du temps d'activité et du nombre de bouchée
    pics = []
    inALoop = False
    x = 0
    activity_time = 0
    debut_pic = 0
    threshold = 5
    for i in range(indice_debut, indice_fin):
        y = df["Ptot"].iloc[i]
        if not inALoop and y + threshold < df["Ptot"].shift(-1).iloc[i]:
            inALoop = True
            debut_pic = df["time"].iloc[i]
            pics.append([debut_pic, df["Ptot"].iloc[i]])
            x = y
        elif inALoop:
            if x > y + threshold:
                fin_pic = df["time"].iloc[i]
                activity_time += fin_pic - debut_pic
                pics[len(pics) - 1] += [fin_pic, df["Ptot"].iloc[i]]
                inALoop = False
                bouche += 1
            elif x > y:
                inALoop = False
                del pics[len(pics) - 1]
    print(f"Le temps d'activité est de : {activity_time} sec")
    ratio = activity_time / temps_repas
    print(f"Le ratio d'activité est : {ratio * 100} %")

    print(f"Le nombre de bouchée pendant le repas est : {bouche}")

    # Création de graphiques avec Plotly

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
    for x in pics:
        fig.add_vline(
            x=x[0],
            line=dict(color="red", width=2), annotation_text=f"{x[1]}", annotation_position="top"
        )
        try:
            fig.add_vline(
                x=x[2],
                line=dict(color="red", width=2), annotation_text=f"{x[3]}", annotation_position="top"
            )
            # Add a shape (rectangle) to fill the area between the vertical lines
            fig.add_shape(
                type="rect",
                x0=x[0],
                x1=x[2],
                y0=0,  # Adjust y0 and y1 as needed
                y1=max_value,  # Adjust y0 and y1 as needed
                fillcolor="rgba(255, 0, 0, 0.2)",  # Transparent red fill color
                line=dict(color="rgba(255, 0, 0, 0)")  # No border line
            )
        except:
            pass

    
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
    fig.add_annotation(
    text=f"Temps total du repas : {tempsmin} minutes et {tempssec} secondes",
    xref="paper",  # Utilise les coordonnées relatives à la zone du graphique
    yref="paper",
    x=0.98,  # Position sur l'axe des x, 1 étant tout à droite
    y=0.94,  # Position sur l'axe des y, 1 étant tout en haut
    showarrow=False,
    font=dict(
        family="Calibri, monospace",
        size=16,
        color="#000000"
    ))

    fig.add_annotation(
    text=f" Masse de repas consommée: {poids_consome} grammes",
    xref="paper",  # Utilise les coordonnées relatives à la zone du graphique
    yref="paper",
    x=0.98,  # Position sur l'axe des x, 1 étant tout à droite
    y=0.90,  # Position sur l'axe des y, 1 étant tout en haut
    showarrow=False,
    font=dict(
        family="Calibri, monospace",
        size=16,
        color="#000000"
    ))

    fig.write_html(filepath)
    bouche = 0 

