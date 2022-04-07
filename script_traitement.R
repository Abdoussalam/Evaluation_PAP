#-------------------------------------------------------------------------------
# Scripts pour l'évaluation des activités du PAP 2022 de l'INS du Niger
# © Abdoussalam ZAKARI Avril 2022
#-------------------------------------------------------------------------------

# Chargement des packages de travail

if(!require(purrr)) install.packages("purrr")
if(!require(dplyr)) install.packages("dplyr")
if(!require(readxl)) install.packages("readxl")
if(!require(openxlsx)) install.packages("openxlsx")
if(!require(here)) install.packages("here")

#-------------------------------------------------------------------------------
# Importation des feuilles excel renseignées par les structures

## Le nom du fichier contenant les données doit respecter le format:
## evaluation_trimestre_annee.xlsx

annee <- 2022          # A adapter au besoin
trimestre <- "t1"      # A adapter selon le trimestre : t1, t2, t3, t4

# Chemin d'accès au fichier des données

nom_fichier <- paste("evaluation",
                     trimestre,
                     annee,
                     sep = "_") %>% 
  paste0(".xlsx")

chemin_data <- here("donnees", nom_fichier)

## Noms des classeurs du fichier excel
list_feuil <- list(chemin_data, 
                sheet = excel_sheets(chemin_data))

## Importation de toutes les feuilles
sous_pap <- pmap(list_feuil, read_excel) 

#-------------------------------------------------------------------------------
# Fonctions pour renommer les colonnes et sélectionner les plus importantes

nom_col <- function(df){
  df %>% 
    rename(realisation = contains("Etat"), 
           structure = contains("tructure"),
           prevision = contains("Prévision"),
           division = contains("ivision"),
           service = contains("ervice"),
           prioritaire = contains("ossier")) 
}

col_select <- function(df){
  df %>% 
    select(structure, 
           division,
           service,
           prevision,
           realisation, 
           prioritaire)
}

#-------------------------------------------------------------------------------
# Fusion de l'ensemble des données en un tableau et sélection des tâches prévues

sous_pap <- sous_pap %>% 
  map(nom_col) %>% 
  map(col_select) %>% 
  reduce(union_all) %>%
  filter(prevision == 1)

#-------------------------------------------------------------------------------
# Score par service, division et direction

## Fonction pour évaluer les taux d'exécution du sous pap
score <- function(df = sous_pap){
  df %>% 
    summarise(`Nombre de taches prévues` = sum(prevision, na.rm = T),
              `Nombre de taches réalisées` = sum(realisation, na.rm = T),
              `Taux de réalisation (%)` = round(100 * `Nombre de taches réalisées` / 
                `Nombre de taches prévues`,1)
    )
    
}

## Ensemble sous PAP
ensemble = score()

## Taux par service
service = sous_pap %>% 
  group_by(structure, division, service) %>% 
  score()

## taux par Division / Unité
division = sous_pap %>% 
  group_by(structure, division) %>% 
  score()


## Taux d'exécution des dossiers prioritaires

dossier_prio <- sous_pap %>% 
  filter(prioritaire == 1) %>% 
  score() %>% 
  mutate(Dossier = "Ensemble") %>% 
  select(4, 1:3)


## Taux par structure

## Conformément au manuel du mécanisme d'incitation à la performance, pour les 
## membres du COMIDIR, le score trimestriel correspond à la moyenne pondérée des 
## taux de réalisation des activités de leurs structures (75%) et
## du taux d'exécution des dossiers prioritaires (25%)

structure = sous_pap %>% 
  group_by(structure) %>% 
  score() %>% 
  
  mutate(`Taux d'exécution dossiers prioritaires (%)` = 
           dossier_prio$`Taux de réalisation (%)`,
         `Score final (75% taux de réalisation + 25% taux dossiers prioritaires)` = 
           round(`Taux de réalisation (%)` * 0.75 + 
           dossier_prio$`Taux de réalisation (%)` * 0.25 , 1)) %>% 
  
  #Ajout d'une ligne ensemble
  add_row(structure = "Ensemble", 
          `Nombre de taches prévues` = ensemble$`Nombre de taches prévues`,
          `Nombre de taches réalisées` = ensemble$`Nombre de taches réalisées`,
          `Taux de réalisation (%)` = ensemble$`Taux de réalisation (%)`,
          `Taux d'exécution dossiers prioritaires (%)` = 
            dossier_prio$`Taux de réalisation (%)`,
          
          `Score final (75% taux de réalisation + 25% taux dossiers prioritaires)` = 
            round(`Taux de réalisation (%)` * 0.75 + 
                    dossier_prio$`Taux de réalisation (%)` * 0.25,1)
            
         )


#-------------------------------------------------------------------------------
## Exportation des résultats vers excel

resultats <- list(Direction = structure,
                  Division = division,
                  Service = service,
                  Dossiers_prioritaires = dossier_prio)

# Fichiers des résultats pour le trimestre 

fichier_res <- paste("resultats", 
                     trimestre, annee, 
                     sep = "_") %>% 
  paste0(".xlsx")

chemin_res <- here("resultats", fichier_res)

write.xlsx(resultats, chemin_res, asTable = T)

  