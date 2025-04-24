import pandas as pd
from Feuille import Feuille

class Test_spe:
    def __init__(self, nom: str, feuille: Feuille):
        self.nom = nom
        self.feuille = feuille

    def val_max(self, val_max: float, colonne: str):
        return self.valider_colonne(
            colonne=colonne,
            condition=lambda x: x <= val_max,
            message=f"strictement <= {val_max}",
            erreur_message=f"valeurs ≤ {val_max}",
            recherche="max"
        )
    
    def val_min(self, val_min: float, colonne: str):
        return self.valider_colonne(
            colonne=colonne,
            condition=lambda x: x >= val_min,
            message=f"strictement >= {val_min}",
            erreur_message=f"valeurs ≥ {val_min}",
            recherche="min"
        )
    
    def val_entre(self, val_min: float, val_max: float, colonne: str):
        return self.valider_colonne(
            colonne=colonne,
            condition=lambda x: (x >= val_min) & (x <= val_max),
            message=f"entre {val_min} et {val_max} (inclus)",
            erreur_message=f"valeurs hors de [{val_min}, {val_max}]",
            recherche="entre"
        )


    def valider_colonne(self, colonne: str, condition, message: str, erreur_message: str, recherche=""):
        erreurs = {}
        message_final = ""
        try:
            df = self.feuille.get_feuille()
        except Exception as e:
            return f"❌ Erreur lors de la lecture du fichier : {e}"
    
        # Recherche les indices correspondant au nom de colonne
        ligne_nom = self.feuille.entete_fin - 1
        ligne_unite = self.feuille.ligne_unite
        ligne_data = self.feuille.data_debut
    
        indices_colonnes = [i for i, val in enumerate(df.iloc[ligne_nom]) if str(val).strip() == colonne]
        if not indices_colonnes:
            return f"⚠️ Colonne '{colonne}' introuvable dans la ligne d'en-tête."
    
        if self.feuille.nb_colonnes_secondaires > 0 and ligne_unite is not None:
            recherche = recherche.lower()
            if recherche in ("min", "max", "entre"):
                sous_colonnes = []
                for i in indices_colonnes:
                    sous_col = str(df.iloc[ligne_unite, i]).lower()
                    if recherche in sous_col:
                        sous_colonnes.append(i)
                if sous_colonnes:
                    indices_colonnes = sous_colonnes
                elif recherche in ("min", "max"):
                    return f"⚠️ Aucune colonne secondaire '{recherche}' détectée pour '{colonne}'."
    
        # Vérifie les données
        for col_index in indices_colonnes:
            valeurs = pd.to_numeric(df.iloc[ligne_data:, col_index], errors='coerce')
            masque = condition(valeurs)
            valeurs_invalides = valeurs[~masque]
            self.feuille.ajouts_erreur(valeurs_invalides.index, col_index)
    
            nom_col_affiche = f"{colonne} (col {col_index})"
            if not valeurs_invalides.empty:
                erreurs[nom_col_affiche] = valeurs_invalides
                message_final += f"❌ Erreurs dans '{nom_col_affiche}' : {len(valeurs_invalides)} valeurs {erreur_message}\n"
            else:
                message_final += f"✅ Toutes les valeurs dans '{nom_col_affiche}' sont {message}\n"
    
        if not erreurs:
            message_final += "\n🎉 Aucune erreur détectée.\n"
        else:
            message_final += "\n🛑 Résumé des erreurs :\n"
            for col, err in erreurs.items():
                message_final += f"- {col} : {len(err)} erreurs\n"
    
        message_final += "======================================================================================= \n"
        return message_final


    def compare_col_fix(self, diff: int, colonne1: str, colonne2: str):
        return self._comparer(colonne1, colonne2, lambda a, b: abs(a - b) > diff, f"diff > {diff}")

    def compare_col_ratio(self, ratio: float, colonne1: str, colonne2: str):
        return self._comparer(colonne1, colonne2, lambda a, b: abs(a - b) > (b * ratio), f"écart > {ratio}x")

    def _comparer(self, colonne1, colonne2, condition, erreur_message):
        try:
            df = self.feuille.get_feuille()
            df_data = df.iloc[self.feuille.data_debut:self.feuille.data_fin or None]

            if self.feuille.ignorer_lignes_vides:
                df_data = df_data.dropna(how='all')

            idx1 = df.iloc[self.feuille.entete_debut].tolist().index(colonne1)
            idx2 = df.iloc[self.feuille.entete_debut].tolist().index(colonne2)

            val1 = pd.to_numeric(df_data.iloc[:, idx1], errors='coerce')
            val2 = pd.to_numeric(df_data.iloc[:, idx2], errors='coerce')

            erreurs = condition(val1, val2)

            lignes_erreurs = erreurs[erreurs].index.tolist()
            self.feuille.ajouts_erreur(lignes_erreurs, idx1, code=2)
            self.feuille.ajouts_erreur(lignes_erreurs, idx2, code=2)

            if lignes_erreurs:
                return f"❌ {len(lignes_erreurs)} erreurs détectées entre '{colonne1}' et '{colonne2}' ({erreur_message})"
            return f"✅ Aucune erreur détectée entre '{colonne1}' et '{colonne2}'."

        except Exception as e:
            return f"Erreur de comparaison : {e}"