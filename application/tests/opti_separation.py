
import pandas as pd

from structure.Fichier import Fichier
from structure.Feuille import Feuille


def split_excel_by_column(feuille :Feuille, column: str, output_file: str):

    input_file= feuille.fichier.chemin 
    sheet_name = feuille.nom
    try:
        # Lire le fichier Excel
        df = pd.read_excel(input_file, sheet_name=sheet_name)

        # Vérifier si la colonne existe
        if column not in df.columns:
            print(f"Valeur de column : {column}|")
            print(f"Valeur de df.columns : {df.columns}")
            
            print(f"Valeur de df.columns.str.strip() : {df.columns.str.strip()}")
            raise ValueError(f"La colonne '{column}' n'existe pas dans la feuille '{sheet_name}'.")

        # Créer un fichier Excel avec un writer
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            # Obtenir les valeurs uniques dans la colonne
            unique_values = df[column].dropna().unique()

            # Créer une feuille par valeur unique
            for value in unique_values:
                # Filtrer les lignes correspondant à la valeur actuelle
                filtered_df = df[df[column] == value]
                # Nom de la feuille limité à 31 caractères
                sheet_name = str(value)[:31]
                # Écrire dans la feuille
                filtered_df.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"Le fichier '{output_file}' a été créé avec succès.")
    except Exception as e:
        print(f"Erreur : {e}")
