import pandas as pd
from fichier import Fichier
from Feuille import Feuille


class Test_spe:
    def __init__(self, nom: str, feuille: Feuille):
        self.nom = nom
        self.feuille = feuille

    def __str__(self):
        return f"Test {self.nom} (fichier: {self.fichier.nom})"

    def val_max(self, val_max: float, colonne: str, min_ou_max: str = None):
        return self.valider_colonne(
            colonne=colonne,
            condition=lambda x: x <= val_max,
            message=f"strictement <= {val_max}",
            erreur_message=f"valeurs ≤ {val_max}",
            min_ou_max=min_ou_max
            )

    def val_min(self, val_min: float, colonne: str, min_ou_max: str = None):
        return self.valider_colonne(
            colonne=colonne,
            condition=lambda x: x >= val_min,
            message=f"strictement >= {val_min}",
            erreur_message=f"valeurs ≥ {val_min}",
            min_ou_max=min_ou_max
        )

    def val_entre(self, val_min: float, val_max: float, colonne: str, min_ou_max: str = None):
        return self.valider_colonne(
            colonne=colonne,
            condition=lambda x: (x >= val_min) & (x <= val_max),
            message=f"entre {val_min} et {val_max} (inclus)",
            erreur_message=f"valeurs hors de [{val_min}, {val_max}]",
            min_ou_max=min_ou_max
        )

    def valider_colonne(self, colonne: str, condition, message: str, erreur_message: str, min_ou_max: str = None):
        erreurs = {}
        message_final = ""
        try:
            df = self.feuille.get_feuille()
        except Exception as e:
            msg_tmp =f"Erreur lors de la lecture du fichier : {e}"
            message_final += msg_tmp
            print(msg_tmp)
            return
    
        # Trouver tous les indices de colonnes correspondant au nom donné
        indices_colonnes = [i for i, val in enumerate(df.iloc[0]) if val == colonne]
    
        if not indices_colonnes:
            msg_tmp = f"⚠️ La colonne '{colonne}' n'existe pas dans la première ligne."
            message_final += msg_tmp

            raise ValueError(msg_tmp)
    
        ligne_data = self.feuille.taille_entete
    
        # Si on veut cibler uniquement les colonnes "min" ou "max"
        if min_ou_max:
            indices_colonnes = [
                i for i in indices_colonnes
                if str(df.iloc[1, i]).strip().lower() == min_ou_max.lower()
            ]
            if not indices_colonnes:
                msg_tmp = f"⚠️ Aucune colonne nommée '{colonne}' avec attribut '{min_ou_max}' trouvée."
                raise ValueError(msg_tmp)
    
        for col_index in indices_colonnes:
            entete_2 = str(df.iloc[1, col_index])
            valeurs = pd.to_numeric(df.loc[ligne_data:, col_index], errors='coerce')
            masque = condition(valeurs)
            valeurs_invalides = valeurs[~masque]
            self.feuille.ajouts_erreur(valeurs_invalides.index,col_index)

            
    
            col_key = f"{colonne} ({entete_2})" if min_ou_max is None else colonne
    
            if not valeurs_invalides.empty:
                erreurs[col_key] = valeurs_invalides
                msg_tmp = f"❌ Erreurs détectées dans la colonne '{col_key}' ({erreur_message}) :"
                message_final += msg_tmp

                print(msg_tmp)
                print(valeurs_invalides)
            else:
                msg_tmp =f"✅ Toutes les valeurs dans la colonne '{col_key}' sont {message}.\n"
                message_final += msg_tmp

                print(msg_tmp)

        if not erreurs:
            msg_tmp ="\n🎉 Aucune erreur trouvée dans les colonnes analysées.\n"
            message_final += msg_tmp

            print(msg_tmp)
        else:
            msg_tmp ="\n🛑 Des erreurs ont été détectées :\n"
            message_final += msg_tmp

            print(msg_tmp)
            for col, err in erreurs.items():
                msg_tmp =f"- Colonne '{col}' : {len(err)} valeurs hors plage.\n"
                message_final += msg_tmp

                print(msg_tmp)
        print("======================================================================================= \n")
        return message_final
    
    
    
    
    def compare_col_fix(self, diff: int, colonne1: str, colonne2: str):
        erreurs = {}
        message_final = ""
        try:
            df = self.feuille.get_feuille()
        except Exception as e:
            msg_tmp = f"Erreur lors de la lecture du fichier : {e}"
            message_final+= msg_tmp
            print(msg_tmp)
            return

        for colonne in [colonne1, colonne2]:
            if colonne not in df.iloc[0, :].values:
                msg_tmp = f"La colonne '{colonne}' n'existe pas dans le fichier."
                message_final+= msg_tmp
                raise ValueError(msg_tmp)

        col_1_index = df.iloc[0, :].tolist().index(colonne1)
        col_2_index = df.iloc[0, :].tolist().index(colonne2)
        ligne_data = self.feuille.taille_entete

        valeurs_col_1 = pd.to_numeric(df.loc[ligne_data:, col_1_index], errors='coerce')
        valeurs_col_2 = pd.to_numeric(df.loc[ligne_data:, col_2_index], errors='coerce')

        ecarts = abs(valeurs_col_1 - valeurs_col_2)
        masque_erreurs = ecarts > diff

        df_erreurs = pd.DataFrame({
            colonne1: valeurs_col_1[masque_erreurs],
            colonne2: valeurs_col_2[masque_erreurs],
            'Écart absolu': ecarts[masque_erreurs]
        })

        if not df_erreurs.empty:
            erreurs[f"{colonne1} vs {colonne2}"] = df_erreurs
            msg_tmp = f"❌ Erreurs détectées entre '{colonne1}' et '{colonne2}' (écart > {diff}) :"
            message_final+= msg_tmp
            print(msg_tmp)
            print(df_erreurs)
            
            # 💾 Ajout dans self.erreurs pour chaque colonne impliquée
            lignes_fautives = df_erreurs.index.tolist()
            self.feuille.ajouts_erreur(lignes_fautives, col_1_index,2)
            self.feuille.ajouts_erreur(lignes_fautives, col_2_index,2)
        else:
            msg_tmp = f"✅ Toutes les valeurs entre '{colonne1}' et '{colonne2}' sont valides.\n"
            message_final+= msg_tmp
            print(msg_tmp)

        if not erreurs:
            msg_tmp = "\n🎉 Aucune erreur trouvée.\n"
            message_final+= msg_tmp
            print(msg_tmp)
        else:
            msg_tmp = "\n🛑 Des erreurs ont été détectées :"
            message_final+= msg_tmp
            print(msg_tmp)
            for col, err in erreurs.items():
                msg_tmp = f"- Comparaison '{col}' : {len(err)} lignes concernées.\n"
                message_final+= msg_tmp
                print(msg_tmp)
        print("======================================================================================= \n")
        return message_final
    
    def compare_col_ratio(self, ratio: float, colonne1: str, colonne2: str):
        erreurs = {}
        message_final = ""
        try:
            df = self.feuille.get_feuille()
        except Exception as e:
            msg_tmp = f"Erreur lors de la lecture du fichier : {e}"
            print(msg_tmp)
            message_final += msg_tmp
            return

        for colonne in [colonne1, colonne2]:
            if colonne not in df.iloc[0, :].values:
                msg_tmp = f"La colonne '{colonne}' n'existe pas dans le fichier."
                message_final += msg_tmp
                raise ValueError(msg_tmp)
                

        col_1_index = df.iloc[0, :].tolist().index(colonne1)
        col_2_index = df.iloc[0, :].tolist().index(colonne2)
        ligne_data = self.feuille.taille_entete

        valeurs_col_1 = pd.to_numeric(df.loc[ligne_data:, col_1_index], errors='coerce')
        valeurs_col_2 = pd.to_numeric(df.loc[ligne_data:, col_2_index], errors='coerce')

        ecarts = abs(valeurs_col_1 - valeurs_col_2)
        ecart_accepte = valeurs_col_2 * ratio
        masque_erreurs = ecarts > ecart_accepte

        df_erreurs = pd.DataFrame({
            colonne1: valeurs_col_1[masque_erreurs],
            colonne2: valeurs_col_2[masque_erreurs],
            'Écart absolu': ecarts[masque_erreurs],
            'Écart accepté': ecart_accepte[masque_erreurs]
        })

        if not df_erreurs.empty:
            erreurs[f"{colonne1} vs {colonne2}"] = df_erreurs
            msg_tmp = f"❌ Erreurs détectées entre '{colonne1}' et '{colonne2}' (écart > {ratio}x {colonne2}) :"
            print(msg_tmp)
            print(df_erreurs)
            # 💾 Ajout dans self.erreurs pour chaque colonne impliquée
            lignes_fautives = df_erreurs.index.tolist()
            self.feuille.ajouts_erreur(lignes_fautives, col_1_index,2)
            self.feuille.ajouts_erreur(lignes_fautives, col_2_index,2)
        else:
            msg_tmp = f"✅ Toutes les valeurs entre '{colonne1}' et '{colonne2}' respectent le ratio.\n"
            message_final += msg_tmp
            print(msg_tmp)

        if not erreurs:
            msg_tmp = "\n🎉 Aucune erreur trouvée.\n"
            message_final += msg_tmp
            print(msg_tmp)
        else:
            msg_tmp = "\n🛑 Des erreurs ont été détectées :"
            message_final += msg_tmp

            print(msg_tmp)
            for col, err in erreurs.items():
                msg_tmp = f"- Comparaison '{col}' : {len(err)} lignes concernées.\n"
                message_final += msg_tmp

                print(msg_tmp)
        print("======================================================================================= \n")
        return message_final
    
    
    
    
    
    
    
    
    
    
    
    
    
    