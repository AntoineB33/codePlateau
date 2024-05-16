# Chargement des différentes library nécessaires 
library(signal) #Transformation des données en fréquences
library(ggplot2) # Outil de création de graphique
library(gridExtra) # Outil de création de graphique
library(plotly) # Outil de création de graphique
library(openxlsx)

# install.packages("signal")
# install.packages("ggplot2")
# install.packages("gridExtra")
# install.packages("plotly")
# install.packages("openxlsx")


dossier <- "c:/Users/comma/Documents/travail/Polytech/stage s8/code/data/Donnees_brutes_csv" # Selectionne le dossier où sont stockées les données 

fichiers <- list.files(path = dossier, pattern = "\\.xlsx$", full.names = TRUE)
# fichiers <- list.files(path = dossier, full.names = TRUE) # Renvoie un vecteur avec tous les fichiers contenus dans le dossier

dossier_graphique <- setwd("c:/Users/comma/Documents/travail/Polytech/stage s8/code/result") # Chemin d'enregistrement des différents graphiques 



Tableau_Final <- data.frame(Duree_Totale = numeric(),
                            Poids_Conso = numeric(),
                            Action = numeric(),
                            Duree_activite_Totale = numeric(),
                            Duree_activite_mean = numeric(),
                            Duree_activite_max = numeric(),
                            Duree_activite_min = numeric(),
                            `Proportion_activite_%` = numeric(),
                            Bouchees = numeric(),
                            Num_fichier = character())

# Parcours des fichiers un à un 
for (fichier in fichiers) {
  
  # Lecture du jeu de données brutes
  df <- read.xlsx(fichier)
  
  
  # Définition du poids minimal de l'assiette 
  plate_weight_min <- 100 # en gramme
  data <- data.frame(time = df$time, Ptot = df$Ptot)
  data <- data[data$Ptot > plate_weight_min,]
  
  # Définition des conditions d'une bouchée 
  min_bite_duration <- 1  # en secondes
  min_bite_weight <- 4    # en grammes 
  print(4)
  
  
  
  # Filtre lissage temps / le plateau prend des données toutes les 0.160s en moyenne 
  int_time <- 0.2 # intervalle de temps min
  indice <- 0
  indices_time <- c(which(cumsum(c(diff(data$time),0)) >= int_time)[1])
  data_true <- as.data.frame(data)
  n <- seq(nrow(data))
  n_true <- n
  data_true$Index <- seq(nrow(data_true))
  
  # Tant que le temps n'est pas supérieur à 0.2s , on supprime la ligne 
  
  while(!is.na(indice)) {
    Cumul_duration <- c(cumsum(abs(diff(data_true$time))),0)
    
    indice <- n_true[which((Cumul_duration) >= int_time)][1]
    
    
    if (!is.na(indice)) {
      n_true <- n_true[which(n_true == indice):length(n_true)]
      indices_time <- c(indices_time,indice)
      n_true <- n_true[-1]  
      data_true <- data_true[which(data_true$Index == n_true[1]):nrow(data_true),]
      
    }
  }
  
  indices_time <- unique(indices_time)
  
  filtered_data <- data[indices_time,]
  
  # Filtre lissage poids / seuil fixé à 4g (min_bite_weight)
  
  rownames(filtered_data) <- seq(nrow(filtered_data))
  seuil_poids <- min_bite_weight
  indice <- 0
  filtered_data_true <- filtered_data
  filtered_data_wo_noise <- filtered_data
  n <- seq(nrow(filtered_data))
  n_true <- n
  filtered_data_true$Index <- seq(nrow(filtered_data_true))
  
  while(!is.na(indice)) {
    val_ini <- as.numeric(filtered_data_true[1,"Ptot"])
    
    indice <- n_true[which(abs(filtered_data_true$Ptot - val_ini) >= seuil_poids)][1]-1
    
    if (!is.na(indice)) {
      n_true <- n_true[which(n_true == indice):length(n_true)]
      indice_commun <- which(filtered_data_true$time[1] == filtered_data_wo_noise$time)
      
      filtered_data_wo_noise$Ptot[indice_commun:indice] <- val_ini
      n_true <- n_true[-1]  
      filtered_data_true <- filtered_data_true[which(filtered_data_true$Index == n_true[1]):nrow(filtered_data_true),]
      
    }
  }
  
  
  filtered_data <- filtered_data_wo_noise
  
  indice_min_duration <-  which(cumsum(c(diff(filtered_data$time),0)) >= min_bite_duration)[1] # Renvoie le nombre de valeur nécessaire pour atteindre la durée minimale d'une bouchée fixée à 1s. Ici le filtre est à 0.2 donc indice_min_duration = 5
  
  
  # Fonction permettant de transformer mes données dans le domaine fréquentiel
  
  butter_bandstop_filter <- function(data, lowcut, highcut, fs, order) {
    nyq <- 0.5 * fs
    low <- lowcut / nyq
    high <- highcut / nyq
    
    # Filtre passe-bas
    b1 <- butter(order, high, type = "low", plane = "z")
    # Filtre passe-haut
    b2 <- butter(order, low, type = "high", plane = "z")
    
    # Appliquer les deux filtres
    y <- signal:::filter(b1, signal:::filter(b2, data, circular = FALSE), circular = FALSE)
    
    return(y)
  }
  
  # Fonction : détection des points d'inflexion et création de segments
  
  find_inflexion_points <- function(data) {
    # Calcul de la dérivée seconde (approximation de la courbure)
    curvature <- abs(diff(diff(data)))
    # Trouver les indices des points d'inflexion
    inflexion_indices <- which(curvature > 5) # Fixé arbitrairement 
    
    return(inflexion_indices)
  }
  
  create_segments <- function(indices) {
    segments <- list()
    current_segment <- c(indices[1])
    
    for (i in 2:length(indices)) {
      if (indices[i] - indices[i-1] <= indice_min_duration) {
        current_segment <- c(current_segment, seq(indices[i-1] + 1, indices[i] - 1))
      } else {
        # Ajouter le segment seulement s'il contient plus de indice_min_duration éléments
        if (length(current_segment) > indice_min_duration) {
          segments <- append(segments, list(unique(sort(current_segment))))
        }
        current_segment <- c(indices[i])
      }
    }
    # Ajouter le dernier segment
    if (length(current_segment) > indice_min_duration) {
      segments <- append(segments, list(current_segment))
    }
    
    return(segments)
  }
  
  complete_segments <- function(segments, max_diff) {
    completed_segments <- list()
    
    for (segment in segments) {
      completed_segment <- segment[1]
      for (i in 2:length(segment)) {
        diff <- segment[i] - segment[i-1]
        if (diff <= max_diff) {
          # Ajouter les indices manquants consécutifs
          missing_indices <- seq(segment[i-1] + 1, segment[i] - 1)
          completed_segment <- c(completed_segment, missing_indices)
        }
        completed_segment <- sort(unique(c(completed_segment, segment[i])))
      }
      completed_segments <- append(completed_segments, list(completed_segment))
    }
    
    return(completed_segments)
  }
  
  # Fonction calculant la durée de chacune des actions
  
  calculate_segment_durations_time <- function(segments, time_data) {
    durations <- sapply(segments, function(segment) {
      if (length(segment) == 1) {
        return(0)
      } else {
        return(time_data[segment[length(segment)]] - time_data[segment[1]])
      }
    })
    return(durations)
  }
  
  # Fonction calculant le poids entre chaque action
  
  calculate_segment_weight <- function(segments, Ptot_data) {
    Ptot_min <- sapply(segments, function(segment) {
      if (length(segment) == 1) {
        return(0) # Si le segment n'a qu'un seul élément, retourne 0
      } else {
        return(min(Ptot_data[segment]))
      }
    })
    return(Ptot_min)
  }
  
  # Fonction pour completer les segments action / non action 
  
  segment_consecutive <- function(indices) {
    segments <- list()
    current_segment <- c(indices[1])
    
    for (i in 2:length(indices)) {
      if (indices[i] - indices[i-1] == 1) {
        current_segment <- c(current_segment, indices[i])
      } else {
        segments <- append(segments, list(current_segment))
        current_segment <- c(indices[i])
      }
    }
    segments <- append(segments, list(current_segment))
    
    return(segments)
  }
  
  # Fonction pour obtenir le dernier indice dans chaque segment
  dernier_indice_segment <- function(segment_action, indice_bites) {
    dernier_indices <- numeric(length(indice_bites))
    
    for (i in seq_along(indice_bites)) {
      segment <- segment_action[[indice_bites[i]]]
      dernier_indices[i] <- segment[length(segment)]
    }
    
    return(dernier_indices)
  }
  
  
  # Filtrage de la série temporelle avec un filtre bandstop
  ts_data_filt <- butter_bandstop_filter(filtered_data$Ptot, 0.5, 1, 1 / mean(diff(filtered_data$time)), 4)
  
  # Détection des points d'inflexion
  inflexion_points <- find_inflexion_points(ts_data_filt)
  
  # Création des segments à partir des points d'inflexion
  segment_action <- create_segments(inflexion_points)
  
  # Compléter les segments pour assurer une durée minimale
  segment_action <- complete_segments(segment_action, indice_min_duration) 
  
  
  
  
  
  # Identification des segments d'action et de non-action
  indice_action <- unlist(segment_action)
  indice_non_action <- setdiff(seq_along(filtered_data$time), indice_action)
  segment_non_action <- segment_consecutive(indice_non_action)
  
  
  
  # A la fin de ça, il y a deux listes, action et non-action, qui contiennent des indices de lognes 
  
  
  
  # Calcul des durées et poids pour les segments d'action
  durations_action_time <- calculate_segment_durations_time(segment_action, filtered_data$time)
  mean_action <- mean(durations_action_time)
  max_action <- max(durations_action_time)
  min_action <- min(durations_action_time)
  duree_totale_action <- sum(durations_action_time)
  # On commence par l'action 2 car la 1 correspond au dépot de l'assiette sur le plateau 
  duree_repas <- abs(filtered_data$time[segment_action[[2]][1]] - filtered_data$time[segment_action[[length(segment_action)]][length(segment_action[[length(segment_action)]])]])  
  nb_action <- length(segment_action)
  proportion_action <- duree_totale_action / duree_repas 
  
  
  # Calcul du poids consommé pendant les périodes de non-action
  weight_non_action <- calculate_segment_weight(segment_non_action, filtered_data$Ptot)
  weight_non_action <- weight_non_action[weight_non_action > plate_weight_min]
  poids_conso <- weight_non_action[1] - min(weight_non_action)
  bites <- length(which(diff(weight_non_action) < 0))
  indice_bites <- which(diff(weight_non_action) < 0) + 1
  
  
  # Creation d'un tableau temporaire
  temp_df <- data.frame(Duree_Totale = duree_repas,
                        Poids_Conso = poids_conso,
                        Action = nb_action,
                        Duree_activite_Totale = round(duree_totale_action,3),
                        Duree_activite_mean = round(mean_action,3),
                        Duree_activite_max = round(max_action,3),
                        Duree_activite_min = round(min_action,3),
                        Proportion_activite = round(proportion_action,3)*100,
                        Bouchees = bites,
                        Num_fichier = gsub("\\.xlsx$", "", basename(fichier)))
  
  
  # Ajout de temp_df à la fin de Tableau_Final
  Tableau_Final <- rbind(Tableau_Final, temp_df) # Fusion des deux tableaux
  
  indice_bites <- dernier_indice_segment(segment_action, indice_bites)
  time_bites <- filtered_data$time[indice_bites]
  weight_bites <- filtered_data$Ptot[indice_bites]
  
  
  # Affichage du graphique
  # Définir les couleurs pour les segments d'action
  colors <- rainbow(n = length(segment_action))
  
  # Créer un graphique interactif avec Plotly
  p <- plot_ly(data = filtered_data, x = ~time, y = ~Ptot, type = 'scatter', mode = 'lines', line = list(color = 'black'), name = 'Données filtrées') %>%
    layout(title = paste0("Repas : ", gsub("\\.csv$", "", basename(fichier))),
           xaxis = list(title = 'Temps'),
           yaxis = list(title = 'Ptot'))
  
  # Ajouter des segments d'action avec une couleur différente
  for (i in seq_along(segment_action)) {
    segment <- segment_action[[i]]
    color <- colors[i %% length(colors) + 1]
    segment_data <- filtered_data[segment, ]
    p <- add_trace(p, data = segment_data, x = ~time, y = ~Ptot, type = 'scatter', mode = 'lines', line = list(color = color), name = paste("Action", i))
  }
  
  # Ajout des lignes verticales
  for (i in 1:length(time_bites)) {
    p <- p %>%
      add_trace(x = c(time_bites[i], time_bites[i]), y = c(min(filtered_data$Ptot), weight_bites[i]),
                xend = c(time_bites[i], time_bites[i]), yend = c(min(filtered_data$Ptot), weight_bites[i]),
                line = list(color = 'green', dash = 'dot'), name = paste("Bouchée n°", i))
  }
  # Ajouter la courbe ts_data_filt
  p <- add_trace(p, x = filtered_data$time, y = ts_data_filt, type = 'scatter', mode = 'lines', line = list(color = 'black'), name = 'Analyse fréquentielle')
  p
  # Enregistrer le widget dans le dossier spécifié
  htmlwidgets::saveWidget(p, file = paste0(dossier_graphique, "/Graph_Repas_", gsub("\\.xlsx$", "", basename(fichier)), ".html"))
  
  print(basename(fichier))
  
}
