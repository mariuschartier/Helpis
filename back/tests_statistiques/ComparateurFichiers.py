from scipy.stats import shapiro, anderson, normaltest, levene, bartlett, ttest_ind, mannwhitneyu
import pandas as pd
import tkinter as tk

from fonctions import to_int
from structure.Fichier import Fichier
from structure.Feuille import Feuille
from structure.Entete import Entete


class ComparateurFichiers:
    """
    Classe pour comparer des fichiers Excel et effectuer des tests statistiques."""
    def __init__(self):
        self.feuille :Feuille = None


    def ajouter_feuille(self, feuille):
        """
        Ajoute une feuille à comparer."""
        self.feuille = feuille


    def collecter_donnees(self, colonne: str):
        """Collecte les données d'une colonne spécifique dans la feuille active.
        - colonne: nom complet de la colonne (chemin)"""
        datas = []

        try:
            # Vérification de la colonne dans l'entête
            indice_colonne = self.feuille.entete.placement_colonne[colonne]
        except KeyError:
            raise KeyError(f"❌ Erreur : La colonne '{colonne}' n'existe pas dans placement_colonne.")

        debut = self.feuille.debut_data
        fin = self.feuille.fin_data

        df = self.feuille.df.iloc[debut:fin+1, :]
        serie = df.iloc[:, indice_colonne]

        if not serie.empty:
            datas.append(serie.reset_index(drop=True))

        if not datas:
            print("❌ Aucune donnée valide à concaténer. Vérifiez les feuilles ajoutées ou les filtres.")
            raise ValueError("❌ Aucune donnée valide à concaténer. Vérifiez les feuilles ajoutées ou les filtres.")

        return pd.concat(datas, ignore_index=True)


    def tester_normalite(self, colonne, methode="shapiro", seuil=0.05):
        """Teste la normalité des valeurs d'une colonne avec la méthode choisie."""
        try:
            serie = pd.to_numeric(self.collecter_donnees(colonne), errors='coerce').dropna()

            if len(serie) < 3:
                return {"stat": None, "p_value": None, "normal": False}

            if methode == "shapiro":
                stat, p = shapiro(serie)
                return {"stat": stat, "p_value": p, "normal": p > seuil}

            elif methode == "dagostino":
                stat, p = normaltest(serie)
                return {"stat": stat, "p_value": p, "normal": p > seuil}

            elif methode == "anderson":
                result = anderson(serie)
                stat = result.statistic
                seuils = result.critical_values
                niveaux = result.significance_level
                index = next(i for i, s in enumerate(niveaux) if s <= seuil * 100)
                normal = stat < seuils[index]
                return {"stat": stat, "p_value": None, "normal": normal}

            else:
                raise ValueError("Méthode inconnue : shapiro, dagostino, anderson")

        except Exception as e:
            print(f"Erreur dans tester_normalite: {e}")
            return {"stat": None, "p_value": None, "normal": False}


    def tester_homogeneite_variances(self, variable: str, groupe: str, methode="levene", seuil=0.05):
        """
        Teste l'homogénéité des variances pour une variable en fonction des groupes.
        - variable: nom complet de la colonne de la variable (chemin)
        - groupe: nom complet de la colonne des groupes (chemin)
        """
        resultats = {}

        try:
            serie_var = self.collecter_donnees(variable)
            serie_grp = self.collecter_donnees(groupe)

            if len(serie_var) != len(serie_grp):
                raise ValueError("Longueurs différentes entre variable et groupe")

            data = pd.DataFrame({variable: serie_var, groupe: serie_grp}).dropna()
            groupes_uniques = data[groupe].unique()

            if len(groupes_uniques) < 2:
                resultats = {"stat": None, "p_value": None, "homogene": False}
            else:
                echantillons = [data[data[groupe] == g][variable] for g in groupes_uniques]
                if methode == "levene":
                    stat, p = levene(*echantillons)
                elif methode == "bartlett":
                    stat, p = bartlett(*echantillons)
                else:
                    raise ValueError("Méthode inconnue : choisir 'levene' ou 'bartlett'")

                resultats = {"stat": stat, "p_value": p, "homogene": p > seuil}

        except Exception as e:
            print(f"Erreur dans tester_homogeneite_variances: {e}")
            resultats = {"stat": None, "p_value": None, "homogene": False}

        return resultats


    def tester_comparaison_groupes(self, variable, groupe, groupe_1: str, groupe_2: str, methode="student", seuil=0.05):
        """
        Compare une variable entre deux groupes avec un test statistique.
        - variable: nom complet de la colonne contenant les valeurs numériques
        - groupe: nom complet de la colonne de regroupement
        """
        resultats = {}

        try:
            # 🔹 Récupération des séries depuis la feuille
            serie_var = self.collecter_donnees(variable)
            serie_grp = self.collecter_donnees(groupe)

            # 🔹 Fusion des deux en DataFrame
            data = pd.DataFrame({variable: serie_var, groupe: serie_grp}).dropna()

            # 🔸 Typage explicite pour comparaison
            if data[groupe].dtype == object:
                groupe_1 = str(groupe_1)
                groupe_2 = str(groupe_2)
            else:
                groupe_1 = pd.to_numeric(groupe_1, errors="coerce")
                groupe_2 = pd.to_numeric(groupe_2, errors="coerce")

            # 🔹 Filtrage des deux groupes
            data = data[data[groupe].isin([groupe_1, groupe_2])]

            g1 = pd.to_numeric(data[data[groupe] == groupe_1][variable], errors="coerce").dropna()
            g2 = pd.to_numeric(data[data[groupe] == groupe_2][variable], errors="coerce").dropna()

            # 🔸 Vérification de validité
            if g1.empty or g2.empty:
                return {
                    "stat": None,
                    "p_value": None,
                    "significatif": False,
                    "error": "Un des groupes est vide ou contient des données non numériques."
                }

            # 🧪 Test statistique
            if methode == "student":
                stat, p = ttest_ind(g1, g2, equal_var=True)
            elif methode == "mannwhitney":
                stat, p = mannwhitneyu(g1, g2)
            else:
                raise ValueError("Méthode non reconnue : choisir 'student' ou 'mannwhitney'.")

            resultats = {
                "groupe_1": str(groupe_1),
                "groupe_2": str(groupe_2),
                "stat": stat,
                "p_value": p,
                "significatif": p < seuil
            }

        except Exception as e:
            resultats = {
                "stat": None,
                "p_value": None,
                "significatif": False,
                "error": f"Erreur : {e}"
            }

        return resultats



    def tester_comparaison_moyennes_hebdo(self, variable, groupe, groupe_1, groupe_2, methode="student", seuil=0.05):
        """
        Compare les moyennes hebdomadaires d'une variable entre deux groupes.
        """
        resultats = {}

        try:
            # 🔹 Récupération des séries depuis la feuille
            serie_var = self.collecter_donnees(variable)
            serie_grp = self.collecter_donnees(groupe)

            # 🔹 Fusion des deux en DataFrame
            data = pd.DataFrame({variable: serie_var, groupe: serie_grp}).dropna()

            # 🔸 Typage explicite pour comparaison
            if data[groupe].dtype == object:
                groupe_1 = str(groupe_1)
                groupe_2 = str(groupe_2)
            else:
                groupe_1 = pd.to_numeric(groupe_1, errors="coerce")
                groupe_2 = pd.to_numeric(groupe_2, errors="coerce")

            # 🔹 Filtrage des deux groupes
            data = data[data[groupe].isin([groupe_1, groupe_2])]

            g1 = pd.to_numeric(data[data[groupe] == groupe_1][variable], errors="coerce").dropna()
            g2 = pd.to_numeric(data[data[groupe] == groupe_2][variable], errors="coerce").dropna()

            # 🔸 Vérification de validité
            if g1.empty or g2.empty:
                return {
                    "stat": None,
                    "p_value": None,
                    "significatif": False,
                    "error": "Un des groupes est vide ou contient des données non numériques."
                }

            # Application du test
            if methode == "student":
                stat, p = ttest_ind(g1, g2, equal_var=True)
            elif methode == "mannwhitney":
                stat, p = mannwhitneyu(g1, g2)
            else:
                raise ValueError("Méthode inconnue : choisir 'student' ou 'mannwhitney'")

            resultats = {
                "groupe_1": str(groupe_1),
                "groupe_2": str(groupe_2),
                "stat": stat,
                "p_value": p,
                "significatif": p < seuil
            }

        except Exception as e:
            resultats = {"stat": None, "p_value": None, "significatif": False, "error": str(e)}

        return resultats
