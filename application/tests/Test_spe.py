
import tests.fonctions as fn
import tkinter as tk
import pandas as pd

from structure.Fichier import Fichier
from structure.Feuille import Feuille


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
            )

    def val_min(self, val_min: float, colonne: str):
        return self.valider_colonne(
            colonne=colonne,
            condition=lambda x: x >= val_min,
            message=f"strictement >= {val_min}",
            erreur_message=f"valeurs ≥ {val_min}",
        )

    def val_entre(self, val_min: float, val_max: float, colonne: str):
        return self.valider_colonne(
            colonne=colonne,
            condition=lambda x: (x >= val_min) & (x <= val_max),
            message=f"entre {val_min} et {val_max} (inclus)",
            erreur_message=f"valeurs hors de [{val_min}, {val_max}]",

        )
    
    def compare_col_fix(self, difference: float, colonne1: str,colonne2: str):
        return self.valider_comparaison(
            colonne1=colonne1,
            colonne2=colonne2,
            condition=lambda x,y: abs(x-y) > difference,
            message=f"Difference < {difference} entre {colonne1} et {colonne2}",
            erreur_message=f"Difference > {difference} entre {colonne1} et {colonne2}",
        )
    
    def compare_col_ratio(self, ratio: float, colonne1: str,colonne2: str):
        print(f"ratio : {ratio}")
        return self.valider_comparaison(
            colonne1=colonne1,
            colonne2=colonne2,
            condition=lambda x,y: abs(x-y) > y*ratio,
            message=f"Difference < {ratio} entre {colonne1} et {colonne2}",
            erreur_message=f"Difference > {ratio} entre {colonne1} et {colonne2}",
        )
    

    def valider_colonne(self, colonne: str, condition, message: str, erreur_message: str):
        erreurs = {}
        message_final = ""
        try:
            df = self.feuille.get_feuille()
        except Exception as e:
            msg_tmp = f"Erreur lors de la lecture du fichier : {e}"
            message_final += msg_tmp
            print(msg_tmp)
            return
        print(self.feuille.entete.placement_colonne)
        # Trouver tous les indices de colonnes correspondant au nom donné
        try:
            indices_colonnes = self.feuille.entete.placement_colonne[colonne]
        except Exception as e:
            msg_tmp = f"⚠️ La colonne '{colonne}' n'existe pas dans la première ligne."
            message_final += msg_tmp
            raise ValueError(msg_tmp)

        ligne_data = self.feuille.debut_data
        print(ligne_data)
        entete_2 = str(df.iloc[1, indices_colonnes])
        valeurs = pd.to_numeric(df.loc[ligne_data:, indices_colonnes], errors='coerce')

        # Vérification et conversion en série si nécessaire
        if isinstance(valeurs, (int, float)):
            valeurs = pd.Series([valeurs])

        masque = condition(valeurs)
        valeurs_invalides = valeurs[~masque]
        self.feuille.ajouts_erreur(valeurs_invalides.index, indices_colonnes)

        col_key = f"{colonne} ({entete_2})"

        if not valeurs_invalides.empty:
            erreurs[col_key] = valeurs_invalides
            msg_tmp = f"❌ Erreurs détectées dans la colonne '{col_key}' ({erreur_message}) :"
            message_final += msg_tmp

            print(msg_tmp)
            print(valeurs_invalides)
        else:
            msg_tmp = f"✅ Toutes les valeurs dans la colonne '{col_key}' sont {message}.\n"
            message_final += msg_tmp

            print(msg_tmp)

        if not erreurs:
            msg_tmp = "\n🎉 Aucune erreur trouvée dans les colonnes analysées.\n"
            message_final += msg_tmp

            print(msg_tmp)
        else:
            msg_tmp = "\n🛑 Des erreurs ont été détectées :\n"
            message_final += msg_tmp

            print(msg_tmp)
            for col, err in erreurs.items():
                msg_tmp = f"- Colonne '{col}' : {len(err)} valeurs hors plage.\n"
                message_final += msg_tmp

                print(msg_tmp)
        print("======================================================================================= \n")
        return message_final
    
    def valider_comparaison(self, colonne1: str,colonne2: str, condition, message: str, erreur_message: str):
        erreurs = {}
        message_final = ""
        try:
            df = self.feuille.get_feuille()
        except Exception as e:
            msg_tmp = f"Erreur lors de la lecture du fichier : {e}"
            message_final += msg_tmp
            print(msg_tmp)
            return
        print(self.feuille.entete.placement_colonne)
        # Trouver tous les indices de colonnes correspondant au nom donné
        try:
            indices_colonne1 = self.feuille.entete.placement_colonne[colonne1]
        except Exception as e:
            msg_tmp = f"⚠️ La colonne '{indices_colonne1}' n'existe pas dans la première ligne."
            message_final += msg_tmp
            raise ValueError(msg_tmp)
        

        try:
            indices_colonne2 = self.feuille.entete.placement_colonne[colonne2]
        except Exception as e:
            msg_tmp = f"⚠️ La colonne '{indices_colonne2}' n'existe pas dans la première ligne."
            message_final += msg_tmp
            raise ValueError(msg_tmp)

        ligne_data = self.feuille.debut_data
        entete_1 = str(df.iloc[1, indices_colonne1])
        entete_2 = str(df.iloc[1, indices_colonne2])

        valeurs1 = pd.to_numeric(df.loc[ligne_data:, indices_colonne1], errors='coerce')
        valeurs2 = pd.to_numeric(df.loc[ligne_data:, indices_colonne2], errors='coerce')

        # Vérification et conversion en série si nécessaire
        if isinstance(valeurs1, (int, float)) and isinstance(valeurs2, (int, float)):
            valeurs1 = pd.Series([valeurs1])
            valeurs2 = pd.Series([valeurs2])

        masque = condition(valeurs1, valeurs2)
        valeurs_invalides1 = valeurs1[~masque]
        valeurs_invalides2 = valeurs2[~masque]

        self.feuille.ajouts_erreur(valeurs_invalides1.index, indices_colonne1,2)
        self.feuille.ajouts_erreur(valeurs_invalides2.index, indices_colonne2,2)
        

        col_key = f"{colonne1} ({entete_1}) et {colonne2} ({entete_2})"

        if not valeurs_invalides1.empty or not valeurs_invalides2.empty:
            erreurs[col_key] = f"{valeurs_invalides1} et {valeurs_invalides2}"
            msg_tmp = f"❌ Erreurs détectées dans la colonne '{col_key}' ({erreur_message}) :"
            message_final += msg_tmp

            print(msg_tmp)
        else:
            msg_tmp = f"✅ Toutes les valeurs dans la colonne '{col_key}' sont {message}.\n"
            message_final += msg_tmp

            print(msg_tmp)

        if not erreurs:
            msg_tmp = "\n🎉 Aucune erreur trouvée dans les colonnes analysées.\n"
            message_final += msg_tmp

            print(msg_tmp)
        else:
            msg_tmp = "\n🛑 Des erreurs ont été détectées :\n"
            message_final += msg_tmp

            print(msg_tmp)
            for col, err in erreurs.items():
                msg_tmp = f"- Colonne '{col}' : {len(err)} valeurs hors plage.\n"
                message_final += msg_tmp

                print(msg_tmp)
        print("======================================================================================= \n")
        return message_final
    
    
    # def compare_col_fix(self, diff: int, colonne1: str, colonne2: str):
    #     erreurs = {}
    #     message_final = ""
    #     try:
    #         df = self.feuille.get_feuille()
    #     except Exception as e:
    #         msg_tmp = f"Erreur lors de la lecture du fichier : {e}"
    #         message_final+= msg_tmp
    #         print(msg_tmp)
    #         return

    #     for colonne in [colonne1, colonne2]:
    #         if colonne not in df.iloc[0, :].values:
    #             msg_tmp = f"La colonne '{colonne}' n'existe pas dans le fichier."
    #             message_final+= msg_tmp
    #             raise ValueError(msg_tmp)

    #     col_1_index = df.iloc[0, :].tolist().index(colonne1)
    #     col_2_index = df.iloc[0, :].tolist().index(colonne2)
    #     ligne_data = self.feuille.debut_data
        
    #     valeurs_col_1 = pd.to_numeric(df.loc[ligne_data:, col_1_index], errors='coerce')
    #     valeurs_col_2 = pd.to_numeric(df.loc[ligne_data:, col_2_index], errors='coerce')

    #     ecarts = abs(valeurs_col_1 - valeurs_col_2)
    #     masque_erreurs = ecarts > diff

    #     df_erreurs = pd.DataFrame({
    #         colonne1: valeurs_col_1[masque_erreurs],
    #         colonne2: valeurs_col_2[masque_erreurs],
    #         'Écart absolu': ecarts[masque_erreurs]
    #     })

    #     if not df_erreurs.empty:
    #         erreurs[f"{colonne1} vs {colonne2}"] = df_erreurs
    #         msg_tmp = f"❌ Erreurs détectées entre '{colonne1}' et '{colonne2}' (écart > {diff}) :"
    #         message_final+= msg_tmp
    #         print(msg_tmp)
    #         print(df_erreurs)
            
    #         # 💾 Ajout dans self.erreurs pour chaque colonne impliquée
    #         lignes_fautives = df_erreurs.index.tolist()
    #         self.feuille.ajouts_erreur(lignes_fautives, col_1_index,2)
    #         self.feuille.ajouts_erreur(lignes_fautives, col_2_index,2)
    #     else:
    #         msg_tmp = f"✅ Toutes les valeurs entre '{colonne1}' et '{colonne2}' sont valides.\n"
    #         message_final+= msg_tmp
    #         print(msg_tmp)

    #     if not erreurs:
    #         msg_tmp = "\n🎉 Aucune erreur trouvée.\n"
    #         message_final+= msg_tmp
    #         print(msg_tmp)
    #     else:
    #         msg_tmp = "\n🛑 Des erreurs ont été détectées :"
    #         message_final+= msg_tmp
    #         print(msg_tmp)
    #         for col, err in erreurs.items():
    #             msg_tmp = f"- Comparaison '{col}' : {len(err)} lignes concernées.\n"
    #             message_final+= msg_tmp
    #             print(msg_tmp)
    #     print("======================================================================================= \n")
    #     return message_final
    
    # def compare_col_ratio(self, ratio: float, colonne1: str, colonne2: str):
    #     erreurs = {}
    #     message_final = ""
    #     try:
    #         df = self.feuille.get_feuille()
    #     except Exception as e:
    #         msg_tmp = f"Erreur lors de la lecture du fichier : {e}"
    #         print(msg_tmp)
    #         message_final += msg_tmp
    #         return

    #     for colonne in [colonne1, colonne2]:
    #         if colonne not in df.iloc[0, :].values:
    #             msg_tmp = f"La colonne '{colonne}' n'existe pas dans le fichier."
    #             message_final += msg_tmp
    #             raise ValueError(msg_tmp)
                

    #     col_1_index = df.iloc[0, :].tolist().index(colonne1)
    #     col_2_index = df.iloc[0, :].tolist().index(colonne2)
    #     ligne_data = self.feuille.entete.taille_entete

    #     valeurs_col_1 = pd.to_numeric(df.loc[ligne_data:, col_1_index], errors='coerce')
    #     valeurs_col_2 = pd.to_numeric(df.loc[ligne_data:, col_2_index], errors='coerce')

    #     ecarts = abs(valeurs_col_1 - valeurs_col_2)
    #     ecart_accepte = valeurs_col_2 * ratio
    #     masque_erreurs = ecarts > ecart_accepte

    #     df_erreurs = pd.DataFrame({
    #         colonne1: valeurs_col_1[masque_erreurs],
    #         colonne2: valeurs_col_2[masque_erreurs],
    #         'Écart absolu': ecarts[masque_erreurs],
    #         'Écart accepté': ecart_accepte[masque_erreurs]
    #     })

    #     if not df_erreurs.empty:
    #         erreurs[f"{colonne1} vs {colonne2}"] = df_erreurs
    #         msg_tmp = f"❌ Erreurs détectées entre '{colonne1}' et '{colonne2}' (écart > {ratio}x {colonne2}) :"
    #         print(msg_tmp)
    #         print(df_erreurs)
    #         # 💾 Ajout dans self.erreurs pour chaque colonne impliquée
    #         lignes_fautives = df_erreurs.index.tolist()
    #         self.feuille.ajouts_erreur(lignes_fautives, col_1_index,2)
    #         self.feuille.ajouts_erreur(lignes_fautives, col_2_index,2)
    #     else:
    #         msg_tmp = f"✅ Toutes les valeurs entre '{colonne1}' et '{colonne2}' respectent le ratio.\n"
    #         message_final += msg_tmp
    #         print(msg_tmp)

    #     if not erreurs:
    #         msg_tmp = "\n🎉 Aucune erreur trouvée.\n"
    #         message_final += msg_tmp
    #         print(msg_tmp)
    #     else:
    #         msg_tmp = "\n🛑 Des erreurs ont été détectées :"
    #         message_final += msg_tmp

    #         print(msg_tmp)
    #         for col, err in erreurs.items():
    #             msg_tmp = f"- Comparaison '{col}' : {len(err)} lignes concernées.\n"
    #             message_final += msg_tmp

    #             print(msg_tmp)
    #     print("======================================================================================= \n")
    #     return message_final
    
    
    
    
    
    
    
    
    
    
    
    
    
    