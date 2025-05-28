import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import scipy.stats as stats
from tkinter import filedialog, messagebox, ttk,simpledialog


from structure.Feuille import Feuille
from structure.Fichier import Fichier

def plot_histogram_normal( indice_col: int, feuille :Feuille):
    """
    Trace un histogramme avec une courbe normale théorique et un Q-Q plot pour une colonne de données provenant d'un fichier Excel.
    
    Args:
    - file_path (str): Chemin du fichier Excel.
    - sheet_name (str): Nom de la feuille contenant les données.
    - column_name (str): Nom de la colonne contenant les données à analyser.
    
    Returns:
    - None
    """
    # Chargement des données depuis le fichier Excel
    try:
        df = feuille.get_feuille()
        start_row = feuille.debut_data
        end_row = feuille.fin_data
        # Vérification
        print(f"Type de df : {type(df)}")
        print(f"Range de lignes : {start_row} à {end_row}")
        print(f"Colonne : {indice_col}")
        donnees = df.iloc[start_row:end_row, indice_col].dropna()
    except Exception as e:
        print(f"Erreur lors du chargement des données : {e}")
        return

    # Tracer l'histogramme
    plt.hist(donnees, bins=30, density=True, alpha=0.6, color='g', label='Données')

    # Calculer la densité de la courbe normale
    mu, sigma = donnees.mean(), donnees.std()
    x = np.linspace(donnees.min(), donnees.max(), 100)
    normal_curve = stats.norm.pdf(x, mu, sigma)

    # Tracer la courbe normale
    plt.plot(x, normal_curve, color='r', linewidth=2, label='Courbe normale théorique')
    plt.xlabel('Valeurs')
    plt.ylabel('Densité')
    plt.title('Histogramme et courbe normale')
    plt.legend()
    plt.show()




def plot_qqplot(indice_col: int, feuille):
    """
    Trace un Q-Q plot pour une colonne de données provenant d'un fichier Excel.
    """
    try:
        df = feuille.get_feuille()
        if not isinstance(df, pd.DataFrame):
            raise TypeError("La feuille n'est pas un DataFrame.")

        start_row = feuille.debut_data
        end_row = feuille.fin_data

        if not (0 <= indice_col < df.shape[1]):
            raise IndexError(f"Indice de colonne {indice_col} hors limites (max {df.shape[1]-1})")
        if end_row > len(df) or start_row < 0 or start_row >= end_row:
            raise ValueError("La plage de lignes est invalide.")

        donnees = df.iloc[start_row:end_row, indice_col].dropna()

        if not isinstance(donnees, pd.Series):
            raise TypeError("Les données sélectionnées ne sont pas une série valide.")
        if donnees.empty:
            raise ValueError("Les données sélectionnées sont vides.")
        if len(donnees) < 3:
            raise ValueError("Le Q-Q plot nécessite au moins 3 valeurs.")
        donnees = [i for  i in donnees]
        plt.figure(figsize=(8, 6))
        stats.probplot(donnees, plot=plt)
        plt.title('Q-Q Plot')
        plt.show()

    except Exception as e:
        from tkinter import messagebox
        messagebox.showerror("Erreur", f"Erreur lors de l'affichage du Q-Q plot : {e}")
        print(f"Erreur : {e}")