from typing import Callable
import pandas as pd
from Feuille import Feuille


class Test_gen:
    def __init__(self, nom: str, critere: list):
        self.nom = nom
        self.critere = critere

    def __str__(self):
        return f"Test {self.nom} (critère {self.critere})"

    def val_max(self, feuille: Feuille, val_max: float):
        return self.valider_colonnes(
            feuille,
            condition=lambda x: x <= val_max,
            message=f"<= {val_max}",
            erreur_message=f"valeurs > {val_max}"
        )

    def val_min(self, feuille: Feuille, val_min: float):
        return self.valider_colonnes(
            feuille,
            condition=lambda x: x >= val_min,
            message=f">= {val_min}",
            erreur_message=f"valeurs < {val_min}"
        )

    def val_entre(self, feuille: Feuille, val_min: float, val_max: float):
        return self.valider_colonnes(
            feuille,
            condition=lambda x: (x >= val_min) & (x <= val_max),
            message=f"entre {val_min} et {val_max} inclus",
            erreur_message=f"valeurs hors [{val_min}, {val_max}]"
        )

    def valider_colonnes(
        self,
        feuille: Feuille,
        condition: Callable[[pd.Series], pd.Series],
        message: str,
        erreur_message: str
    ):
        erreurs = {}
        df = feuille.get_feuille()
        ligne_unite = feuille.ligne_unite
        lignes_data = df.iloc[feuille.data_debut : feuille.data_fin or None]

        message_final = ""

        for col in df.columns:
            nom_col = str(df.loc[ligne_unite, col])

            if nom_col in self.critere:
                valeurs = pd.to_numeric(lignes_data[col], errors='coerce')

                if feuille.ignorer_lignes_vides:
                    valeurs = valeurs.dropna()

                masques_valides = condition(valeurs)
                valeurs_invalides = valeurs[~masques_valides]

                if not valeurs_invalides.empty:
                    erreurs[col] = valeurs_invalides
                    feuille.ajouts_erreur(valeurs_invalides.index + feuille.data_debut, col)
                    message_final += f"❌ {len(valeurs_invalides)} erreurs dans colonne {col} ({erreur_message})\n"
                else:
                    message_final += f"✅ Colonne {col} : toutes les valeurs sont {message}\n"

        if not erreurs:
            message_final += f"\n🎉 Aucune erreur dans les colonnes marquées : {self.critere}\n"
        else:
            message_final += "\n🛑 Erreurs détectées dans :\n"
            for col, err in erreurs.items():
                message_final += f"- Colonne {col} : {len(err)} erreurs\n"

        message_final += "=======================================================================================\n"
        return message_final
