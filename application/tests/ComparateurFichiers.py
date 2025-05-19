from scipy.stats import shapiro, anderson, normaltest, levene, bartlett, ttest_ind, mannwhitneyu
import pandas as pd
import tkinter as tk
from tests.fonctions import to_int

class ComparateurFichiers:
    def __init__(self):
        self.feuilles = []


    def ajouter_feuille(self, feuille):
        self.feuilles.append(feuille)


    def collecter_donnees(self):
        datas = []
        for feuille in self.feuilles:
            debut = feuille.debut_data
            fin = feuille.fin_data
            df = feuille.df.iloc[debut:fin+1, :].copy()
            if not df.empty:
                datas.append(df.reset_index(drop=True))
        if not datas:
            raise ValueError("Aucune donnée valide à concaténer. Vérifiez les feuilles ajoutées.")
        return pd.concat(datas, ignore_index=True)

    def tester_normalite(self, methode="shapiro", seuil=0.05):
        """Teste la normalité colonne par colonne avec la méthode choisie."""
        df = self.collecter_donnees()
        resultats = {}
        print (df)

        for colonne in df.columns:
            try:
                serie = pd.to_numeric(df[colonne], errors='coerce').dropna()
                if len(serie) < 3:
                    resultats[colonne] = {"stat": None, "p_value": None, "normal": False}
                    continue

                if methode == "shapiro":
                    stat, p = shapiro(serie)
                    normal = p > seuil
                    resultats[colonne] = {"stat": stat, "p_value": p, "normal": normal}

                elif methode == "dagostino":
                    stat, p = normaltest(serie)
                    normal = p > seuil
                    resultats[colonne] = {"stat": stat, "p_value": p, "normal": normal}

                elif methode == "anderson":
                    result = anderson(serie)
                    stat = result.statistic
                    seuils = result.critical_values
                    significance_levels = result.significance_level
                    seuil_index = next(i for i, s in enumerate(significance_levels) if s <= seuil*100)
                    seuil_critique = seuils[seuil_index]
                    normal = stat < seuil_critique
                    resultats[colonne] = {"stat": stat, "p_value": None, "normal": normal}
                else:
                    raise ValueError("Méthode inconnue : choisir 'shapiro', 'dagostino' ou 'anderson'")

            except Exception as e:
                resultats[colonne] = {"stat": None, "p_value": None, "normal": False}

        return resultats




    def tester_homogeneite_variances(self, colonnes_groupes, methode="levene", seuil=0.05):
        """
        Teste l'homogénéité des variances pour chaque variable entre les groupes.
        colonnes_groupes : dict -> {"variable" : "groupe"}
        méthode : "levene" (par défaut) ou "bartlett"
        """
        df = self.collecter_donnees()
        resultats = {}
        print(colonnes_groupes)
        for variable, groupe in colonnes_groupes.items(): 
            try:
                data = df[[variable, groupe]].dropna()
                groupes_uniques = data[groupe].unique()
                print(data,groupes_uniques)

                if len(groupes_uniques) < 2:
                    resultats[variable] = {"stat": None, "p_value": None, "homogene": False}
                    continue

                echantillons = [data[data[groupe] == g][variable] for g in groupes_uniques]
                if methode == "levene":
                    stat, p = levene(*echantillons)
                elif methode == "bartlett":
                    stat, p = bartlett(*echantillons)
                else:
                    raise ValueError("Méthode inconnue : choisir 'levene' ou 'bartlett'")

                resultats[variable] = {"stat": stat, "p_value": p, "homogene": p > seuil}
                print(stat)
            except Exception as e:
                resultats[variable] = {"stat": None, "p_value": None, "homogene": False}

        return resultats


    def tester_comparaison_groupes(self, variable, groupe, groupe_1: str, groupe_2: str, methode="student", seuil=0.05):
        """
        Compare une variable entre deux groupes avec un test statistique.
        """
        df = self.collecter_donnees()
        resultats = {}

        try:
            data = df[[variable, groupe]].dropna()

            # Assurer le bon typage pour filtrer les groupes
            if data[groupe].dtype == object:
                groupe_1 = str(groupe_1)
                groupe_2 = str(groupe_2)
            else:
                groupe_1 = pd.to_numeric(groupe_1, errors="coerce")
                groupe_2 = pd.to_numeric(groupe_2, errors="coerce")

            # Filtrage des deux groupes
            data = data[data[groupe].isin([groupe_1, groupe_2])]

            # Extraire les échantillons numériques
            g1 = pd.to_numeric(data[data[groupe] == groupe_1][variable], errors="coerce").dropna()
            g2 = pd.to_numeric(data[data[groupe] == groupe_2][variable], errors="coerce").dropna()

            # Vérification de la présence de données dans les deux groupes
            if g1.empty or g2.empty:
                return {
                    "stat": None,
                    "p_value": None,
                    "significatif": False,
                    "error": "Un des groupes est vide ou contient des données non numériques."
                }

            # Test statistique
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
                "error": str(e)
            }

        return resultats


    def tester_comparaison_moyennes_hebdo(self, variable, groupe, groupe_1, groupe_2, methode="student", seuil=0.05):
        """
        Compare les moyennes hebdomadaires d'une variable entre deux groupes.
        """
        df = self.collecter_donnees()
        resultats = {}

        try:
            # Nettoyage de base
            data = df[[variable, groupe]].dropna()

            # Forcer la compatibilité des types
            if data[groupe].dtype == object:
                groupe_1 = str(groupe_1)
                groupe_2 = str(groupe_2)
            else:
                groupe_1 = pd.to_numeric(groupe_1, errors="coerce")
                groupe_2 = pd.to_numeric(groupe_2, errors="coerce")

            data = data[data[groupe].isin([groupe_1, groupe_2])]

            # Extraction des groupes avec conversion en float
            g1 = pd.to_numeric(data[data[groupe] == groupe_1][variable], errors="coerce").dropna()
            g2 = pd.to_numeric(data[data[groupe] == groupe_2][variable], errors="coerce").dropna()

            # Vérification
            if g1.empty or g2.empty:
                return {"stat": None, "p_value": None, "significatif": False, "error": "Un des groupes est vide ou absent."}

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
