from structure.Fichier import Fichier
from structure.Feuille import Feuille

import pandas as pd
from typing import Optional, Callable


class Test_gen:
    def __init__(self, nom: str, critere: list):
        self.nom = nom
        self.critere = critere

    def __str__(self):
        return f"Test {self.nom} (critère {self.critere})"

    def val_max(self, feuille: Feuille, val_max: float):
        # Vérifie que les valeurs sont >= val_max
        return self.valider_colonnes(
            feuille,
            condition=lambda x: x <= val_max,
            message=f"strictement <= {val_max}",
            erreur_message=f"valeurs ≤ {val_max}"
        )

    def val_min(self, feuille: Feuille, val_min: float):
        # Vérifie que les valeurs sont <= val_min
        return self.valider_colonnes(
            feuille,
            condition=lambda x: x >= val_min,
            message=f"strictement >= {val_min}",
            erreur_message=f"valeurs ≥ {val_min}"
        )

    def val_entre(self, feuille: Feuille, val_min: float, val_max: float):
        # Vérifie que les valeurs sont entre val_min et val_max
        return self.valider_colonnes(
            feuille,
            condition=lambda x: (x >= val_min) & (x <= val_max),
            message=f"entre {val_min} et {val_max} (inclus)",
            erreur_message=f"valeurs hors de [{val_min}, {val_max}]"
        )

    def valider_colonnes(
        self,
        feuille: Feuille,
        condition: Callable[[pd.Series], pd.Series],
        message: str,                  # ce message sert à donner un contexte "valeurs entre", etc.
        erreur_message: str):
        erreurs = {}
        df = feuille.get_feuille()
        ligne_symbole = feuille.entete.ligne_unite
        ligne_data = feuille.debut_data
        colonne_symbole = []
        for cle in feuille.entete.placement_colonne:
            for  crit in self.critere:
                if crit in cle:
                    colonne_symbole.append(feuille.entete.placement_colonne[cle])

        print(ligne_data)
        message_final = ""  
        
        if colonne_symbole == []:   
            print(f"Erreur : le critère '{self.critere}' n'existe pas dans le DataFrame.")
            message_final = f"Le critère '{self.critere}' n'existe pas dans le DataFrame.\n"
            return message_final
                
        print(ligne_data)
        message_final = ""  
    
        for col in df.columns:
            nom_col = str(df.loc[ligne_symbole, col])
            if  col  in colonne_symbole:

                msg_tmp = f"📊 Colonne '{col+1}' détectée comme {nom_col} ({self.critere}).\n"
                print(msg_tmp)
                # message_final += msg_tmp
                
                valeurs = pd.to_numeric(df[col].iloc[ligne_data:], errors='coerce')
                masques_valides = condition(valeurs)
                valeurs_invalides = valeurs[~masques_valides]
    
                if not valeurs_invalides.empty:
                    erreurs[col] = valeurs_invalides
                    feuille.ajouts_erreur(valeurs_invalides.index, col)
                    msg_tmp = f"❌ Erreurs détectées dans la colonne {col+1} ({erreur_message})\n"
                    # message_final += msg_tmp
                    print(msg_tmp)
                else:
                    msg_tmp = f"✅ Toutes les valeurs dans la colonne {col+1} sont {message}.\n"
                    # message_final += msg_tmp
                    print(msg_tmp)
    
        if not erreurs:
            msg_tmp = f"\n🎉 Aucune erreur trouvée dans les colonnes marquées en {self.critere}.\n"
            print(msg_tmp)
            message_final += msg_tmp
        else:
            message_final += "\n🛑 Des erreurs ont été détectées dans les colonnes suivantes :\n"
            for col, err in erreurs.items():
                msg_tmp = f"- Colonne {col+1} : {len(err)} valeurs hors plage.\n"
                print(msg_tmp)
                message_final += msg_tmp
    
        print( "======================================================================================= \n")
        return message_final
    
        
        
        
        
        
        
    
    
    
    
    
    
    
    
    
    
    
    
    