import pandas as pd
import re



# --------- PARAMÈTRE : NOM DU FICHIER EXCEL ---------
fichier_excel = "Mesures.xlsx"  # <-- remplace par le nom réel de ton fichier
def recherche_erreur(path_file,taille_entete = 2):
    # Charger le fichier Excel
    df = pd.read_excel(fichier_excel, engine="openpyxl")
    
    print("Aperçu des 5 premières lignes :\n", df.head(), "\n")
    
    # --------- INFORMATIONS GÉNÉRALES ---------
    print("📊 Structure du fichier :")
    print(df.info(), "\n")
    
    # --------- VALEURS MANQUANTES ---------
    print("🔍 Valeurs manquantes par colonne :")
    print(df.isnull().sum(), "\n")
    
    # --------- COLONNES VIDES ---------
    colonnes_vides = df.columns[df.isnull().all()]
    if not colonnes_vides.empty:
        print("⚠️ Colonnes entièrement vides :", list(colonnes_vides), "\n")
    
    # --------- DOUBLONS ---------
    doublons = df[df.duplicated()]
    print(f"📛 Nombre de doublons : {len(doublons)}")
    if not doublons.empty:
        print("Exemples de doublons :\n", doublons.head(), "\n")
    
    # --------- DÉTECTION D'ERREURS PAR COLONNE ---------
    # if 'email' in df.columns:
    #     print("📬 Vérification des emails :")
    #     regex_email = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b'
    #     emails_invalides = df[~df['email'].astype(str).str.match(regex_email)]
    #     if not emails_invalides.empty:
    #         print("❌ Emails invalides détectés :")
    #         print(emails_invalides[['email']], "\n")
    
    # ----------- ANALYSE DES COLONNES POUR '%' À LA LIGNE 2 -----------
    erreurs = {}
    
    for col in df.columns:
        entete_2e_ligne = str(df.loc[taille_entete-1, col])
        
        if '%' in entete_2e_ligne:
            print(f"📊 Colonne {col} détectée comme pourcentage (%).")
            
            valeurs = pd.to_numeric(df.loc[taille_entete:, col], errors='coerce')
            valeurs_invalides = valeurs[(valeurs < 0) | (valeurs > 100)]
            
            if not valeurs_invalides.empty:
                erreurs[col] = valeurs_invalides
                print(f"❌ Erreurs détectées dans la colonne {col} (valeurs hors de [0, 100]) :")
                print(valeurs_invalides)
            else:
                print(f"✅ Toutes les valeurs dans la colonne {col} sont valides (entre 0 et 100).\n")
    
    if not erreurs:
        print("\n🎉 Aucune erreur trouvée dans les colonnes marquées en %.\n")
    else:
        print("\n🛑 Des erreurs ont été détectées dans les colonnes suivantes :")
        for col, err in erreurs.items():
            print(f"- Colonne {col} : {len(err)} valeurs hors plage.\n")
    print("======================================================================================= \n")
    
    # ----------- ANALYSE DES COLONNES POUR '° 'ou '°C' À LA LIGNE 2 -----------
    erreurs = {}
    
    for col in df.columns:
        entete_2e_ligne = str(df.loc[taille_entete-1, col])
        
        if '°' in entete_2e_ligne or '°C' in entete_2e_ligne:
            print(f"📊 Colonne {col} détectée comme températue (°C).")
            
            valeurs = pd.to_numeric(df.loc[taille_entete:, col], errors='coerce')
            valeurs_invalides = valeurs[(valeurs < -10) | (valeurs > 50)]
            
            if not valeurs_invalides.empty:
                erreurs[col] = valeurs_invalides
                print(f"❌ Erreurs détectées dans la colonne {col} (valeurs hors de [-10, 50]) :")
                print(valeurs_invalides)
            else:
                print(f"✅ Toutes les valeurs dans la colonne {col} sont valides [-10, 50].\n")
    
    if not erreurs:
        print("\n🎉 Aucune erreur trouvée dans les colonnes marquées en °C.\n")
    else:
        print("\n🛑 Des erreurs ont été détectées dans les colonnes suivantes :")
        for col, err in erreurs.items():
            print(f"- Colonne {col} : {len(err)} valeurs hors plage.\n")
    print("======================================================================================= \n")
    
    # ----------- ANALYSE DES COLONNES POUR 'g' À LA LIGNE 2 -----------
    erreurs = {}
    
    for col in df.columns:
        entete_2e_ligne = str(df.loc[taille_entete-1, col])
        
        if 'g' in entete_2e_ligne :
            print(f"📊 Colonne {col} détectée comme poids (g).")
            
            valeurs = pd.to_numeric(df.loc[taille_entete:, col], errors='coerce')
            valeurs_invalides = valeurs[(valeurs < 0) | (valeurs > 5000)]
            
            if not valeurs_invalides.empty:
                erreurs[col] = valeurs_invalides
                print(f"❌ Erreurs détectées dans la colonne {col} (valeurs hors de [0, 5000]) :")
                print(valeurs_invalides)
            else:
                print(f"✅ Toutes les valeurs dans la colonne {col} sont valides [0, 5000].\n")
    
    if not erreurs:
        print("\n🎉 Aucune erreur trouvée dans les colonnes marquées en g.\n")
    else:
        print("\n🛑 Des erreurs ont été détectées dans les colonnes suivantes :")
        for col, err in erreurs.items():
            print(f"- Colonne {col} : {len(err)} valeurs hors plage.\n")
    print("======================================================================================= \n")
    
    
    
    
    # if 'date_naissance' in df.columns:
    #     print("📅 Vérification des dates de naissance :")
    #     df['date_naissance'] = pd.to_datetime(df['date_naissance'], errors='coerce')
    #     dates_invalides = df[df['date_naissance'].isnull()]
    #     if not dates_invalides.empty:
    #         print("❌ Dates de naissance invalides :")
    #         print(dates_invalides[['date_naissance']], "\n")
    
    
    # --------- FIN ---------
    print("✅ Vérification terminée.")

recherche_erreur(fichier_excel)