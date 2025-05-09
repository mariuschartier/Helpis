import openpyxl
from structure.Fichier import Fichier
from structure.Feuille import Feuille
from datetime import datetime
from openpyxl import Workbook
import pandas as pd

import subprocess
import sys
# import pkg_resources

def process_excel_data(input_file='results/Ponte Salle 1.xlsx', 
                       sheet_name='Données de ponte', 
                       output_file='results/processed_data.xlsx'):
    # Charger le fichier Excel avec les valeurs calculées
    wb = openpyxl.load_workbook(input_file, data_only=True)

    # Vérifier si la feuille existe
    if sheet_name not in wb.sheetnames:
        print(f"La feuille '{sheet_name}' n'existe pas dans le fichier.")
        return

    ws = wb[sheet_name]
    data = []

    # Lire toutes les lignes de données tout en conservant None pour les cellules vides
    for row in ws.iter_rows(min_row=1, values_only=True):
        # Ajouter chaque ligne comme une liste, tout en gérant les None
        data.append(list(row))

    # Traitement et gestion des None
    nb_colonnes = len(data[0]) if data else 0  # Nombre de colonnes de la première ligne
    structured_data = []

    for r_idx, row in enumerate(data, 1):
        structured_row = []
        for c_idx in range(nb_colonnes):
            if len(row) <= c_idx:
                cell = None
                value_to_set = structured_row[c_idx - 1] if c_idx > 0 else None
            else:
                cell = row[c_idx]
                if cell is None:
                    value_to_set = structured_row[c_idx - 1] if c_idx > 0 else None
                else:
                    value_to_set = cell

            structured_row.append(value_to_set)

        structured_data.append(structured_row)

    # Créer un nouveau classeur et une nouvelle feuille
    new_wb = openpyxl.Workbook()
    new_ws = new_wb.active
    new_ws.title = "Données traitées"

    # Écrire les données structurées dans le fichier
    for r_idx, row in enumerate(structured_data, 1):
        for c_idx, value in enumerate(row, 1):
            new_ws.cell(row=r_idx, column=c_idx, value=value)

    # Enregistrer le fichier Excel
    new_wb.save(output_file)
    print(f"Les données ont été enregistrées dans '{output_file}'.")

    # Afficher les 100 premières lignes de la structure finale
    for index, row in enumerate(structured_data[:100]):  # Limiter à 100 lignes
        print(f"Ligne {index + 1}: {row}")  # Afficher chaque ligne

# Exemple d'utilisation
# process_excel_data()
    
    
    
    
# # Exemple d'utilisation
# input_file = 'results/Ponte Salle 1.xlsx'
# output_file = 'output.xlsx'
# sheet_name = "Données de ponte" # Remplacez par le nom de votre feuille
# process_xlsx(input_file, output_file, sheet_name)



def determine_jour(date_str):
    # Parser la date
    dt = datetime.strptime(date_str, "%d/%m/%y %H:%M")

    # Extraire jour du mois et mois en format 2 chiffres
    jour_mois = dt.strftime("%d")  # exemple : '02'
    mois = dt.strftime("%m")       # exemple : '06'

    return jour_mois, mois

def moyenne_par_jour(feuille,output_file ,date_col=0):
    wb = Workbook()
    ws = wb.active
    ws.title = "moyenne"

    nb_lignes = feuille.nb_ligne
    nb_colonnes = min(feuille.nb_colonne, 21)

    # Copier les lignes d'entête
    for row_idx in range(feuille.taille_entete):
        for col_idx in range(nb_colonnes):
            valeur = feuille.df.iloc[row_idx, col_idx]
            ws.cell(row=row_idx + 1, column=col_idx + 1, value=valeur)

    # Regrouper les lignes par jour
    jours_dict = {}
    for idx in range(feuille.taille_entete, nb_lignes):
        date_cell = feuille.df.iloc[idx, date_col]
        if pd.isna(date_cell):
            continue
        if isinstance(date_cell, datetime):
            dt = date_cell
        else:
            try:
                dt = datetime.strptime(str(date_cell), "%d/%m/%y %H:%M")
            except Exception:
                continue  # Ignorer si la cellule n'est pas une date valide

        jour, mois = dt.strftime("%d"), dt.strftime("%m")
        clef_jour = f"{jour}/{mois}"
        clef_date = datetime(2000, int(mois), int(jour))  # On met 2000 comme annee fictive

        if clef_jour not in jours_dict:
            jours_dict[clef_jour] = {"rows": [], "date": clef_date}
        jours_dict[clef_jour]["rows"].append(idx)

    # Calculer la moyenne pour chaque jour
    ligne_resultat = feuille.taille_entete + 1
    for clef_jour, info in sorted(jours_dict.items(), key=lambda x: x[1]["date"]):
        rows = info["rows"]
        moyennes = []
        for col_index in range(nb_colonnes):
            somme = 0
            count = 0
            for row in rows:
                val = feuille.df.iloc[row, col_index]
                if isinstance(val, (int, float)):
                    somme += val
                    count += 1
            moyenne = somme / count if count > 0 else None
            moyennes.append(moyenne)

        # écrire le jour en 1re colonne
        ws.cell(row=ligne_resultat, column=1, value=clef_jour)
        # écrire les moyennes à partir de la 2ème colonne
        for idx, moyenne in enumerate(moyennes, start=1):
            ws.cell(row=ligne_resultat, column=idx, value=moyenne)

        ligne_resultat += 1

    # Enregistrer le fichier Excel
    wb.save(output_file)
    print(f"Les données ont été enregistrées dans '{output_file}'.")



from datetime import datetime
from openpyxl import Workbook
import pandas as pd

def determine_semaine(date_str):
    # Parser la date
    dt = datetime.strptime(date_str, "%d/%m/%y %H:%M")

    # Extraire annee et semaine ISO
    annee, semaine, _ = dt.isocalendar()
    return f"{annee}-S{semaine:02d}", dt

def moyenne_par_semaine(feuille, output_file,date_col=0):
    wb = Workbook()
    ws = wb.active
    ws.title = "moyenne"

    nb_lignes = feuille.nb_ligne
    nb_colonnes = min(feuille.nb_colonne, 21)

    # Copier les lignes d'entête
    for row_idx in range(feuille.taille_entete):
        for col_idx in range(nb_colonnes):
            valeur = feuille.df.iloc[row_idx, col_idx]
            ws.cell(row=row_idx + 1, column=col_idx + 1, value=valeur)

    # Regrouper les lignes par semaine
    semaines_dict = {}
    for idx in range(feuille.taille_entete, nb_lignes):
        date_cell = feuille.df.iloc[idx, date_col]
        if pd.isna(date_cell):
            continue
        if isinstance(date_cell, datetime):
            dt = date_cell
        else:
            try:
                dt = datetime.strptime(str(date_cell), "%d/%m/%y %H:%M")
            except Exception:
                continue  # Ignorer si la cellule n'est pas une date valide

        clef_semaine, dt_obj = determine_semaine(dt.strftime("%d/%m/%y %H:%M"))

        if clef_semaine not in semaines_dict:
            semaines_dict[clef_semaine] = {"rows": [], "date": dt_obj}
        semaines_dict[clef_semaine]["rows"].append(idx)

    # Calculer la moyenne pour chaque semaine
    ligne_resultat = feuille.taille_entete + 1
    for clef_semaine, info in sorted(semaines_dict.items(), key=lambda x: x[1]["date"]):
        rows = info["rows"]
        moyennes = []
        for col_index in range(nb_colonnes):
            somme = 0
            count = 0
            for row in rows:
                val = feuille.df.iloc[row, col_index]
                if isinstance(val, (int, float)):
                    somme += val
                    count += 1
            moyenne = somme / count if count > 0 else None
            moyennes.append(moyenne)

        # écrire la semaine en 1re colonne
        ws.cell(row=ligne_resultat, column=1, value=clef_semaine)
        # écrire les moyennes à partir de la 2ème colonne
        for idx, moyenne in enumerate(moyennes, start=1):
            ws.cell(row=ligne_resultat, column=idx, value=moyenne)

        ligne_resultat += 1

    # Enregistrer le fichier Excel
    wb.save(output_file)
    print(f"Les données ont été enregistrées dans '{output_file}'.")