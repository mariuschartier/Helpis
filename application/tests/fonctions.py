
from structure.Fichier import Fichier
from structure.Feuille import Feuille
from datetime import datetime
from openpyxl import Workbook
import pandas as pd

import subprocess
import sys
# import pkg_resources

def install_and_import(package, alias=None):
    """Installe un package Python et l'importe sous un alias spécifié. 
    Si le package est déjà installé, il l'importe directement.

    Args:
        package (str): Le nom du package à installer et importer.
        alias (str, optional): L'alias sous lequel importer le package.
    """
    try:
        # Tente d'importer le package
        mod = __import__(package)
        print(f"{package} est déjà installé.")
    except ImportError:
        print(f"{package} n'est pas installé. Installation en cours...")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
        # Réessaie d'importer le package après l'installation
        mod = __import__(package)

    if alias:
        # Utilise globals() pour définir l'alias
        globals()[alias] = mod
        print(f"{package} importé sous l'alias '{alias}'.")
        return mod
        


def determine_jour(date_str):
    # Parser la date
    dt = datetime.strptime(date_str, "%d/%m/%y %H:%M")

    # Extraire jour du mois et mois en format 2 chiffres
    jour_mois = dt.strftime("%d")  # exemple : '02'
    mois = dt.strftime("%m")       # exemple : '06'

    return jour_mois, mois

def moyenne_par_jour(feuille, date_col=0):
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

    wb.save("moyenne_par_jour.xlsx")
    print("Fichier moyenne_par_jour.xlsx sauvegardé.")


from datetime import datetime
from openpyxl import Workbook
import pandas as pd

def determine_semaine(date_str):
    # Parser la date
    dt = datetime.strptime(date_str, "%d/%m/%y %H:%M")

    # Extraire annee et semaine ISO
    annee, semaine, _ = dt.isocalendar()
    return f"{annee}-S{semaine:02d}", dt

def moyenne_par_semaine(feuille, date_col=0):
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

    wb.save("moyenne_par_semaine.xlsx")
    print("Fichier moyenne_par_semaine.xlsx sauvegardé.")


# # # Exemple d'utilisation
# if __name__ == "__main__":
#     jour, mois = determine_jour("02/06/23 15:15")
#     print(f"Jour: {jour}, Mois: {mois}")
#     f1 = Fichier("results/Mesures.xlsx")
#     f1_1 = Feuille(f1,"opti",2)
#     moyenne_par_semaine(f1_1)
#     # Ici, tu peux appeler moyenne_par_jour(sheet) où sheet est une feuille openpyxl
def to_int(val):
    try:
            val = int(val)
    except ValueError:
        print("val is not convertible to an integer")
    return val