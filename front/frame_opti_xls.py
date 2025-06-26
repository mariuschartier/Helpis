from back.manipulation import opti_fichier  # ton module de conversion
from tkinter import filedialog, messagebox, ttk
import tkinter as tk
import ttkbootstrap as ttkb
from ttkbootstrap.constants import *
from pathlib import Path
from back.manipulation import opti_xlsx 
from back.manipulation import opti_separation
import pandas as pd
from fonctions import is_file_locked
from structure.Fichier import Fichier
from structure.Feuille import Feuille
from structure.Selection_col import Selection_col

import os
import sys
from structure.Entete import Entete

class opti_xls(ttkb.Frame):
    """
    Page de manipulation des fichiers Excel (.xls et .xlsx).
    Permet de convertir, améliorer, calculer des moyennes et séparer les données.
    """
    def __init__(self, parent, controller):
        """
        Initialise la page avec les éléments nécessaires pour manipuler les fichiers Excel."""
        super().__init__(parent)
        self.controller = controller

        self.is_xlsx = None
        self.fichier_path = None
        self.df = None
        self.feuille_nom = tk.StringVar()

        self.details_structure = {
            "entete_debut": 0,
            "entete_fin": 0,
            "data_debut": 1,
            "data_fin": None,
            "nb_colonnes_secondaires": 0,
            "ligne_unite": 0,
            "ignorer_vide": True
        }

        self.prepare_dossiers()
        
        self.create_file_frame()
        self.champs_xls()
        self.champs_xlsx()
        self.champs_separation()

        self.desactivation_bouton()

        self.status_label = tk.Label(self, text="", bg="#f4f4f4", fg="green")
        self.status_label.pack(pady=5)
        self.create_excel_preview_frame()
    
# Champ de chargement du fichier et de l'entete =========================================================================================================


    
    def create_file_frame(self):
        """Crée le cadre pour charger le fichier Excel et configurer l'en-tête avec wrapping dynamique et taille minimale."""
        self.file_frame = ttkb.LabelFrame(self, text="1. Charger un fichier Excel")
        self.file_frame.pack(fill="x", expand=False, padx=10, pady=5)

        self.taille_entete_var = tk.StringVar()
        self.taille_entete_var.set("1")
        self.widgets_file_frame = []

        # Widgets à placer dynamiquement
        self.fichier_entry = tk.Entry(self.file_frame, width=60)
        self.widgets_file_frame.append(self.fichier_entry)

        parcourir_btn = ttkb.Button(self.file_frame, text="Parcourir", command=self.controller.bind_button(self.choisir_fichier), width=15)
        self.widgets_file_frame.append(parcourir_btn)

        self.feuille_combo = ttk.Combobox(self.file_frame, textvariable=self.feuille_nom, state="readonly", width=20)
        self.feuille_combo.bind("<<ComboboxSelected>>", lambda e: self.on_feuille_change())
        self.widgets_file_frame.append(self.feuille_combo)

        # Création d'un sous-frame pour aligner label_entete et taille_entete_entry
        entete_frame = tk.Frame(self.file_frame, bg="#f4f4f4")
        label_entete = tk.Label(entete_frame, text="Taille de l'en-tête :")
        label_entete.pack(side="left")

        self.taille_entete_entry = tk.Entry(entete_frame, width=5, textvariable=self.taille_entete_var)
        self.taille_entete_var.trace_add("write", self.on_taille_entete_change)
        self.taille_entete_entry.bind("<KeyRelease>", self.on_key_release_int)
        self.taille_entete_entry.pack(side="left", padx=5)

        self.widgets_file_frame.append(entete_frame)


        self.detail_btn = ttkb.Button(self.file_frame, text="detail", command=self.ouvrir_popup_manipulation, width=10)
        self.widgets_file_frame.append(self.detail_btn)

        self.aide_btn = ttkb.Button(self.file_frame, text="❓ Aide", command=self.ouvrir_aide, width=10)
        self.widgets_file_frame.append(self.aide_btn)



        self.file_frame.bind("<Configure>", lambda event: self.arrange_widgets_file_frame(self.file_frame, self.widgets_file_frame))

        return self.file_frame

    def arrange_widgets_file_frame(self, container, widgets, event=None):
        """ Organise les widgets dans le cadre de chargement du fichier Excel en fonction de la largeur disponible."""
        container.update_idletasks()
        width = container.winfo_width()
        widget_width = 150  # largeur minimale estimée par widget
        num_columns = max(1, width // widget_width)
        # print(f"width = {width}")
        # print(f"widget_width = {widget_width}")
        # print(f"nb_colonne = {num_columns}")

        for widget in container.winfo_children():
            widget.grid_forget()

        for index, widget in enumerate(widgets):
            row = index // num_columns
            col = index % num_columns
            widget.grid(row=row, column=col, padx=5, pady=5, sticky="ew")

        for col in range(num_columns):
            container.grid_columnconfigure(col, weight=1)

    def choisir_fichier(self):
        """
        Ouvre une boîte de dialogue pour sélectionner un fichier Excel (.xls ou .xlsx)."""
        try:
            dossier_data = Path("sauvegardes/results")
            dossier_data.mkdir(parents=True, exist_ok=True)  # Crée le dossier s’il n’existe pas
            
            filepath = filedialog.askopenfilename(
                filetypes=[("Fichiers Excel", "*.xls;*.xlsx")],
                initialdir=dossier_data,  # Dossier par défaut
                title="Choisir un fichier"
            )

            if filepath:
                self.activation_bouton(filepath)
                self.on_feuille_change()
                self.afficher_excel()
     
        except Exception as e:
            print(f"erreur lors de la lecture du fichier: {e}")

    def activation_bouton(self,filepath):
        """Active les boutons et met à jour l'interface en fonction du fichier sélectionné."""
        self.fichier_path = filepath
        self.fichier_entry.delete(0, tk.END)
        self.fichier_entry.insert(0, filepath)

        # Déterminer le type de fichier
        is_xlsx = filepath.endswith(".xlsx")
        is_xls = filepath.endswith(".xls")

        if is_xlsx:
            self.btn_ameliorer.config(state="normal")
            self.btn_moy_jour.config(state="normal")
            self.btn_moy_semaine.config(state="normal")
            self.btn_separation.config(state="normal")
            self.btn_entete_une_ligne.config(state="normal")
            self.detail_btn.config(state="normal")
            self.taille_entete_entry.config(state="normal")

            self.excel_preview_frame.pack(fill="both", expand=True, padx=10, pady=5)
            self.table.grid(row=0, column=0, sticky="nsew")


            self.btn_convertir.config(state="disabled")

        elif is_xls:
            self.btn_convertir.config(state="normal")

            self.btn_ameliorer.config(state="disabled")
            self.btn_moy_jour.config(state="disabled")
            self.btn_moy_semaine.config(state="disabled")
            self.btn_separation.config(state="disabled")
            self.btn_entete_une_ligne.config(state="disabled")
            self.detail_btn.config(state="disabled")
            self.taille_entete_entry.config(state="disabled")

            self.excel_preview_frame.pack_forget()
            self.table.grid_forget()

        else:
            self.btn_convertir.config(state="disabled")
            self.btn_ameliorer.config(state="disabled")
            self.btn_moy_jour.config(state="disabled")
            self.btn_moy_semaine.config(state="disabled")
            self.btn_separation.config(state="disabled")
            self.btn_entete_une_ligne.config(state="disabled")

            self.excel_preview_frame.pack_forget()
            self.table.grid_forget()
            

        if not is_xls:
            try:
                # Sélectionner le moteur approprié
                engine = 'openpyxl' if is_xlsx else None

                try:
                    with pd.ExcelFile(filepath, engine=engine) as xls:
                        feuilles = xls.sheet_names
                except ValueError as e:
                    if not is_xlsx:
                        engine = 'openpyxl'
                        try:
                            with pd.ExcelFile(filepath, engine=engine) as xls:
                                feuilles = xls.sheet_names
                        except Exception as e_openpyxl:
                            raise Exception(f"Erreur avec openpyxl : {e_openpyxl}")
                    else:
                        raise Exception(f"Erreur avec xlrd : {e}")

                self.feuille_combo['values'] = feuilles
                if feuilles:
                    self.feuille_nom.set(feuilles[0])

                print("Le fichier est au format .xlsx" if is_xlsx else "Le fichier est au format .xls")

            except Exception as e:
                messagebox.showerror("Erreur", f"Impossible de lire les feuilles du fichier :\n{e}")

    def desactivation_bouton(self):
        """Désactive tous les boutons et champs de saisie."""
        #entete
        self.detail_btn.config(state="disabled")
        self.taille_entete_entry.config(state="disabled")

        # xls
        self.btn_convertir.config(state="disabled")
        #xlsx
        self.btn_ameliorer.config(state="disabled")
        self.btn_moy_jour.config(state="disabled")
        self.btn_moy_semaine.config(state="disabled")
        self.btn_entete_une_ligne.config(state="disabled")
        # separation
        self.btn_separation.config(state="disabled")

    def ouvrir_aide(self):
        """
        Ouvre une fenêtre d'aide avec des instructions sur l'utilisation de l'application.
        """
        aide_popup = tk.Toplevel(self)
        aide_popup.title("Aide - Utilisation")
        aide_popup.geometry("600x400")

        texte = tk.Text(aide_popup, wrap="word", font=("Segoe UI", 10))
        texte.pack(fill="both", expand=True, padx=10, pady=10)

        contenu = (
                "🔍 Bienvenue sur la page de Manipulation Excel\n\n"
                "Voici comment utiliser l'application :\n"
                "1 - Cliquez sur le bouton 'Parcourir' pour sélectionner un fichier Excel (.xls ou .xlsx).\n"
                "2 - Choisissez la feuille à manipuler dans la liste déroulante.\n"
                "3 - Indiquez la taille de l’en-tête (nombre de lignes au début du tableau).\n"
                "4 - Sélectionnez l'action à effectuer sur le fichier.\n"
                "5 - Choisissez la feuille de sortie si nécessaire.\n\n"
                "Actions disponibles :\n"
                "   • Convertir un fichier .xls en .xlsx\n"
                "   • Améliorer un fichier .xlsx (formatage, nettoyage)\n"
                "   • Calculer la moyenne par jour\n"
                "   • Calculer la moyenne par semaine\n"
                "   • Séparer les données en fonction d'une colonne choisie\n"
                "   • Créer un fichier avec l'en-tête en une seule ligne\n"
        )
        
        texte.insert(tk.END, contenu)

    def on_taille_entete_change(self, *args):
        """
        Met à jour la fin de l'en-tête 
        """

        # Mettre à jour la fin de l'en-tête
        self.details_structure["entete_fin"] = (
            int(self.taille_entete_entry.get()) + self.details_structure["entete_debut"] - 1
            if self.taille_entete_entry.get().isdigit()
            else 0
        )
        self.details_structure["ligne_unite"] = self.details_structure["entete_fin"]
        self.details_structure["data_debut"] = self.details_structure["entete_fin"]+1
        self.enlever_toutes_couleurs()

        self.colorier_lignes_range(
            self.details_structure["entete_debut"],
            self.details_structure["entete_fin"])

        self.dico_entete()     

    def on_feuille_change(self, event=None):
        """Met à jour la feuille sélectionnée et charge les données dans un DataFrame."""
        self.feuille_nom.set(self.feuille_combo.get())
        self.df = pd.read_excel(self.fichier_path, sheet_name=self.feuille_nom.get(), header=None).copy()

        # print(f"Feuille changée : {self.feuille_nom.get()}")
        # print(f"DataFrame shape : {self.df.shape}")
        
        self.details_structure = {
            "entete_debut": 0,
            "entete_fin": 0,
            "data_debut": 1,
            "data_fin": self.df.shape[0] if self.df is not None else None,
            "nb_colonnes_secondaires": 0,
            "ligne_unite": 0,
            "ignorer_vide": True
        }
        self.taille_entete_entry.delete(0, tk.END)
        self.taille_entete_entry.insert(0, str(1))
        print(is_file_locked(self.fichier_path))    

    # Ouvrir le popup de manipulation de l'entete detaillée
    def ouvrir_popup_manipulation(self):
        """Ouvre un popup pour configurer les paramètres avancés de la feuille."""

        # print("details_structure :")
        # print(self.details_structure)

        if self.df is None:
            messagebox.showerror("Erreur", "Un fichier doit être sélectionné.")
            return

        popup = tk.Toplevel(self)
        popup.title("Paramètres avancés de la feuille")
        popup.configure(bg="#ffffff")
        popup.grab_set()

        tk.Label(popup, text="Paramètres de lecture du fichier", font=("Segoe UI", 11, "bold"), bg="#f4f4f4").pack(pady=10)

        champs = [
            ("Début de l'en-tête :", "entete_debut"),
            ("Fin de l'en-tête :", "entete_fin"),
            ("Début des données :", "data_debut"),
            ("Fin des données :", "data_fin"),
            ("Colonnes secondaires :", "nb_colonnes_secondaires"),
            ("Ligne des unités :", "ligne_unite"),
        ]

        entries = {}
        valeurs_par_defaut = {
            "entete_debut": 0,
            "entete_fin": 0,
            "data_debut": 1,
            "data_fin": self.df.shape[0],
            "nb_colonnes_secondaires": 0,
            "ligne_unite": 0,
            "ignorer_vide": True
        }
        data  = self.details_structure if hasattr(self, "details_structure") else valeurs_par_defaut
        

        for label, key in champs:
            frame = tk.Frame(popup, bg="#f4f4f4")
            frame.pack(fill="x", padx=10, pady=2)
            tk.Label(frame, text=label, width=25, anchor="w", bg="#f4f4f4").pack(side="left")

            vcmd = (self.register(lambda val: val.isdigit() or val == ""), '%P')
            entry = tk.Entry(frame, validate="key", validatecommand=vcmd)
            entry.pack(side="left", fill="x", expand=True)

            if key == "data_fin" and data["data_fin"] is None:
                try:
                    valeur_defaut = str(self.df.shape[0])
                except AttributeError:
                    valeur_defaut = ""
            else:
                valeur_defaut = data.get(key, "")

            entry.insert(0, str(valeur_defaut))
            entries[key] = entry

        # ✅ Check : ignorer lignes vides
        ignore_lignes_vides = tk.BooleanVar(value=True)
        frame_cb = tk.Frame(popup, bg="#f4f4f4")
        frame_cb.pack(padx=10, pady=5, anchor="w")

        tk.Checkbutton(frame_cb, text="Ignorer les lignes vides", variable=ignore_lignes_vides, bg="#f4f4f4").pack(side="left")

        def reset_valeur():
            """Réinitialise les valeurs des champs à leurs valeurs par défaut."""
            for key, entry in entries.items():
                if key == "data_fin":
                    try:
                        entry.delete(0, tk.END)
                        entry.insert(0, str(self.df.shape[0]))
                    except AttributeError:
                        entry.delete(0, tk.END)
                        entry.insert(0, "")
                else:
                    valeur_defaut = valeurs_par_defaut.get(key, "")
                    entry.delete(0, tk.END)
                    entry.insert(0, str(valeur_defaut))

            ignore_lignes_vides.set(True)

        ttkb.Button(frame_cb, text="Réinitialisation", command=reset_valeur).pack(side="left", padx=10)

        # ⚠️ Zone de message d'erreur
        label_erreur = tk.Label(popup, text="", fg="red", bg="#f4f4f4", font=("Segoe UI", 9, "italic"))
        label_erreur.pack(pady=5)

        # ✅ Boutons
        frame_btns = tk.Frame(popup, bg="#f4f4f4")
        frame_btns.pack(pady=10)

        def appliquer_parametres():
            try:
                valeurs = {k: int(e.get()) for k, e in entries.items()}
            except ValueError:
                messagebox.showerror("Erreur", "Tous les champs doivent être remplis avec des entiers valides.")
                return

            # Validation
            taille_entete = valeurs["entete_fin"] - valeurs["entete_debut"] + 1
            if taille_entete <= 0:
                messagebox.showerror("Erreur", "L'entête doit contenir au moins une ligne.")
                return

            if valeurs["entete_fin"] >= valeurs["data_debut"]:
                messagebox.showerror("Erreur", "La fin de l'entête doit être avant le début des données.")
                return

            if valeurs["nb_colonnes_secondaires"] >= taille_entete:
                messagebox.showerror("Erreur", "Le nombre de colonnes secondaires doit être inférieur à la taille de l'entête.")
                return

            if not (valeurs["entete_debut"] <= valeurs["ligne_unite"] <= valeurs["entete_fin"]):
                messagebox.showerror("Erreur", "La ligne d'unité doit être comprise dans l'entête.")
                return

            # Appliquer
            valeurs["ignorer_lignes_vides"] = ignore_lignes_vides.get()
            self.details_structure = valeurs
            # print("valeur :")
            # print(valeurs)
            # print("details_structure :")
            # print(self.details_structure)

            if hasattr(self, "taille_entete_entry"):
                self.taille_entete_entry.delete(0, tk.END)
                self.taille_entete_entry.insert(0, str(taille_entete))

            popup.destroy()

        ttkb.Button(frame_btns, text="✅ Appliquer", command=appliquer_parametres).pack(side="left", padx=10)
        ttkb.Button(frame_btns, text="❌ Annuler", command=popup.destroy).pack(side="left", padx=10)




# Champs de manipulation des fichiers xls et xlsx =========================================================================================================
    def champs_xls(self):
        """
        Crée les champs pour manipuler les fichiers Excel (.xls)."""
        frame_action = ttkb.LabelFrame(self, text="2. action sur le fichier xls")
        frame_action.pack(fill="x", padx=10, pady=5)
    
        self.btn_convertir = ttkb.Button(frame_action, text="Convertir en .xlsx", command=self.controller.bind_button(self.convertir_fichier))
        self.btn_convertir.pack(side="left", padx=5)

    def champs_xlsx(self):
        """"Crée les champs pour manipuler les fichiers Excel (.xlsx)."""
        frame_action = ttkb.LabelFrame(self, text="3. action sur le fichier xlsx")
        frame_action.pack(fill="x", padx=10, pady=5)

        self.btn_ameliorer = ttkb.Button(frame_action, text="ameliorer le .xlsx", command=self.controller.bind_button(self.ameliorer_fichier_xlsx))
        self.btn_ameliorer.pack(side="left", padx=5)

        self.btn_moy_jour = ttkb.Button(frame_action, text="moyenne par jour", command=self.controller.bind_button(self.moyenne_par_jour))
        self.btn_moy_jour.pack(side="left", padx=5)

        self.btn_moy_semaine = ttkb.Button(frame_action, text="moyenne par semaine", command=self.controller.bind_button(self.moyenne_par_semaine))
        self.btn_moy_semaine.pack(side="left", padx=5)
        self.btn_entete_une_ligne = ttkb.Button(frame_action, text="Entete en une ligne", command=self.controller.bind_button(self.entete_une_ligne))
        self.btn_entete_une_ligne.pack(side="left", padx=5)

    def champs_separation(self):
        """Crée les champs pour séparer les données d'un fichier Excel par une colonne choisie."""
        frame_action = ttkb.LabelFrame(self, text="4. Création de ficier séparer (xlsx)")
        frame_action.pack(fill="x", padx=10, pady=5)
        self.btn_separation = ttkb.Button(frame_action, text="séparation valeur dans la colonne", command=self.controller.bind_button(self.split_excel_by_column))
        self.btn_separation.pack(side="left", padx=5)

        
# Préparation des dossiers de sauvegarde et de résultats =========================================================================================================
    def prepare_dossiers(self):
        """Prépare les dossiers nécessaires pour les sauvegardes et les résultats."""
        # Récupère le répertoire de l'exécutable
        if hasattr(sys, '_MEIPASS'):
            base_dir = Path(sys._MEIPASS)
        else:
            base_dir = Path(__file__).parent

        sauvegardes_dir = base_dir / 'sauvegardes'

        # Créer les dossiers
        (sauvegardes_dir / 'sauvegardes_tests').mkdir(parents=True, exist_ok=True)
        (sauvegardes_dir / 'results').mkdir(parents=True, exist_ok=True)
        (sauvegardes_dir / 'data').mkdir(parents=True, exist_ok=True)

# Validation de la taille de l'en-tête =========================================================================================================
    def on_key_release_int(self, event):
        """Valide l'entrée de la taille de l'en-tête pour s'assurer qu'elle est un entier positif."""
        if not self.taille_entete_entry.get().isdigit() and self.taille_entete_entry.get() != "":
            messagebox.showwarning("Validation", "Veuillez entrer un nombre entier.")
            self.taille_entete_entry.delete(0, tk.END)


# Construction du dictionnaire d'en-tête =========================================================================================================
    def dico_entete(self):
        """Construit un dictionnaire représentant la structure de l'en-tête du fichier Excel."""
        self.dico_structure = {}
        ligne_entete_debut = self.details_structure.get("entete_debut", 0)
        ligne_entete_fin = self.details_structure.get("entete_fin", 1)
        if self.df is None:
            messagebox.showerror("Erreur", "Un fichier doit être sélectionné.")
            return
        try:
            
            for col_idx in range(len(self.df.columns)):
                current_level = self.dico_structure

                for row_idx in range(ligne_entete_debut, ligne_entete_fin + 1):
                    cell_value = self.df.iloc[row_idx, col_idx]
                    if pd.isna(cell_value):
                        continue
                    cell_value = str(cell_value).strip()

                    if cell_value not in current_level:
                        current_level[cell_value] = {}

                    current_level = current_level[cell_value]

            return self.dico_structure

        except Exception as e:
            messagebox.showerror("Erreur", "Fichier et taille d'entete requis.")
            print(e)
            # messagebox.showerror("Erreur", f"Impossible de construire le dictionnaire d'en-tête : {e}")
            return {}
   

# Methodes de conversion et d'amélioration =========================================================================================================
# convertir un fichier .xls en .xlsx
    def convertir_fichier(self):
        """
        Convertit un fichier Excel (.xls) en format moderne (.xlsx)."""
        if not self.fichier_path or not self.fichier_path.endswith(".xls"):
            messagebox.showerror("Erreur", "Veuillez d'abord sélectionner un fichier .xls valide.")
            return
    
        fichier_destination = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Fichiers Excel modernes", "*.xlsx")],
            title="Enregistrer le fichier converti",
            initialdir="sauvegardes/results",
            initialfile=opti_fichier.fichier_du_chemin(self.fichier_path)
        )
        if not fichier_destination:
            return
    
        try:
            opti_fichier.convertir(self.fichier_path, fichier_destination)
            self.status_label.config(text="✅ Conversion terminée avec succès", fg="green")
            messagebox.showinfo("Succès", f"Fichier converti avec succès :\n{fichier_destination}")
        except Exception as e:
            self.status_label.config(text="❌ Échec de la conversion", fg="red")
            messagebox.showerror("Erreur", f"Erreur lors de la conversion : {e}")


# ameliorer un fichier .xlsx
    def ameliorer_fichier_xlsx(self):
        """
        Améliore un fichier Excel (.xlsx) en le formatant et en le nettoyant."""
        if not self.fichier_path or not self.fichier_path.endswith(".xlsx"):
            messagebox.showerror("Erreur", "Veuillez d'abord sélectionner un fichier .xlsx valide.")
            return
    
        fichier_destination = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Fichiers Excel modernes", "*.xlsx")],
            title="Enregistrer le fichier converti",
            initialdir="sauvegardes/results",
            initialfile=opti_fichier.fichier_du_chemin(self.fichier_path)
        )
        if not fichier_destination:
            return
    
        try:
            opti_xlsx.process_and_format_excel(self.fichier_path,self.feuille_nom.get(), fichier_destination)
            self.status_label.config(text="✅ Conversion terminée avec succès", fg="green")
            messagebox.showinfo("Succès", f"Fichier converti avec succès :\n{fichier_destination}")
        except Exception as e:
            self.status_label.config(text="❌ Échec de la conversion", fg="red")
            messagebox.showerror("Erreur", f"Erreur lors la conversion : {e}")


# Calculer la moyenne par jour
    def moyenne_par_jour(self):
        """Calcule la moyenne des données par jour dans un fichier Excel (.xlsx)."""
        if not self.fichier_path or not self.fichier_path.endswith(".xlsx"):
            messagebox.showerror("Erreur", "Veuillez d'abord sélectionner un fichier .xlsx valide.")
            return
    
        fichier_destination = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Fichiers Excel modernes", "*.xlsx")],
            title="Enregistrer le fichier converti",
            initialdir="sauvegardes/results",
            initialfile=opti_fichier.fichier_du_chemin(self.fichier_path)
        )
        if not fichier_destination:
            return
        fichier = Fichier(self.fichier_path)
        feuille = Feuille(fichier, self.feuille_nom.get(),
                          self.details_structure["data_debut"],
                          self.details_structure["data_fin"],)
        entete = Entete(feuille,self.details_structure["entete_debut"],
                        self.details_structure["entete_fin"],
                        self.details_structure["nb_colonnes_secondaires"],
                        self.details_structure["ligne_unite"],
                        self.dico_entete()
                        )
        feuille.entete = entete
        colonne =  self.afficher_colonne_popup(feuille)
        print(f"Valeur de colonne : {colonne}|")
    
        try:
            num_colonne = feuille.entete.placement_colonne[colonne] 
            f1 = Fichier(self.fichier_path)
            f1_1 = Feuille(f1,self.feuille_nom.get())
            opti_xlsx.moyenne_par_jour(f1_1,fichier_destination,num_colonne)
            self.status_label.config(text="✅ creation terminée avec succès", fg="green")
            messagebox.showinfo("Succès", f"Fichier creer avec succès :\n{fichier_destination}")
        except Exception as e:
            self.status_label.config(text="❌ Échec de la conversion", fg="red")
            messagebox.showerror("Erreur", f"Erreur lors de la creation : {e}")


# Calculer la moyenne par semaine
    def moyenne_par_semaine(self):
        """Calcule la moyenne des données par semaine dans un fichier Excel (.xlsx)."""
        if not self.fichier_path or not self.fichier_path.endswith(".xlsx"):
            messagebox.showerror("Erreur", "Veuillez d'abord sélectionner un fichier .xlsx valide.")
            return
    
        fichier_destination = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Fichiers Excel modernes", "*.xlsx")],
            title="Enregistrer le fichier converti",
            initialdir="sauvegardes/results",
            initialfile=opti_fichier.fichier_du_chemin(self.fichier_path)
        )
        if not fichier_destination:
            return
        fichier = Fichier(self.fichier_path)
        feuille = Feuille(fichier, self.feuille_nom.get(),
                          self.details_structure["data_debut"],
                          self.details_structure["data_fin"],)
        entete = Entete(feuille,self.details_structure["entete_debut"],
                        self.details_structure["entete_fin"],
                        self.details_structure["nb_colonnes_secondaires"],
                        self.details_structure["ligne_unite"],
                        self.dico_entete()
                        )
        feuille.entete = entete
        colonne =  self.afficher_colonne_popup(feuille)
        print(f"Valeur de colonne : {colonne}|")
    
        try:
            num_colonne = feuille.entete.placement_colonne[colonne] 
            f1 = Fichier(self.fichier_path)
            f1_1 = Feuille(f1,self.feuille_nom.get())
            opti_xlsx.moyenne_par_semaine(f1_1,fichier_destination)
            self.status_label.config(text="✅ creation terminée avec succès", fg="green")
            messagebox.showinfo("Succès", f"Fichier creer avec succès :\n{fichier_destination}")
        except Exception as e:
            self.status_label.config(text="❌ Échec de la conversion", fg="red")
            messagebox.showerror("Erreur", f"Erreur lors de la creation : {e}")


# Séparer un fichier Excel par une colonne choisie
    def split_excel_by_column(self):
        """Sépare un fichier Excel (.xlsx) en plusieurs fichiers basés sur les valeurs d'une colonne choisie."""
        fichier = Fichier(self.fichier_path)
        feuille = Feuille(fichier, self.feuille_nom.get(),
                          self.details_structure["data_debut"],
                          self.details_structure["data_fin"],)
        entete = Entete(feuille,self.details_structure["entete_debut"],
                        self.details_structure["entete_fin"],
                        self.details_structure["nb_colonnes_secondaires"],
                        self.details_structure["ligne_unite"],
                        self.dico_entete()
                        )
        feuille.entete = entete
        colonne =  self.afficher_colonne_popup(feuille)
        print(f"Valeur de colonne : {colonne}|")
        if not colonne:
            return
        if not self.fichier_path or not self.fichier_path.endswith(".xlsx"):
            messagebox.showerror("Erreur", "Veuillez d'abord sélectionner un fichier .xlsx valide.")
            return
        fichier_destination = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Fichiers Excel modernes", "*.xlsx")],
            title="Enregistrer le fichier converti",
            initialdir="sauvegardes/results",
            initialfile=opti_fichier.fichier_du_chemin(self.fichier_path)
        )

        if not fichier_destination:
            return
    
        try:
            num_colonne = feuille.entete.placement_colonne[colonne] 
            opti_separation.split_excel_by_column(feuille,num_colonne, fichier_destination)
            self.status_label.config(text="✅ creation terminée avec succès", fg="green")
            messagebox.showinfo("Succès", f"Fichier creer avec succès :\n{fichier_destination}")
        except Exception as e:
            self.status_label.config(text="❌ Échec de la conversion", fg="red")
            message = f"Erreur lors de la creation : {e}"
            print(message)
            messagebox.showerror("Erreur", f"Erreur lors de la creation : {e}")

    def afficher_colonne_popup(self, feuille:Feuille, event=None):
        feuille_obj = feuille
        if not feuille_obj:
            messagebox.showerror("Erreur", "Feuille non trouvée.")
            return None  # ou une valeur par défaut

        # Variable pour stocker la sélection
        self.chemin_selectionne = None

        # Créer la fenêtre popup
        popup = tk.Toplevel(self)
        popup.title("Choisissez la colonne")
        popup.geometry("400x300")
        popup.grab_set()

        tk.Label(popup, text="Choisissez la colonne pour la séparation").pack(pady=10)

        select_colonne = Selection_col(feuille_obj.entete.structure)
        get_path = select_colonne.get_frame_selection_pack(popup)
        select_colonne.pack()
        def valider():
            chemin = get_path()
            chemin = select_colonne.chemin
            if not chemin:
                messagebox.showerror("Erreur", "Veuillez sélectionner une colonne cible.")
                return
            self.chemin_selectionne = chemin
            popup.destroy()

        btn_ok = ttkb.Button(popup, text="valider", command=valider)
        btn_ok.pack(pady=10)

        self.wait_window(popup)  # Attend que la fenêtre soit fermée

        # Après fermeture, retourner la valeur sélectionnée
        return self.chemin_selectionne

    def select_column_path(self, popup,feuille : Feuille):
        """
        Crée une interface pour sélectionner une colonne et ses sous-catégories, et renvoie le chemin sélectionné.

        Args:
        - structure (dict): Dictionnaire représentant la structure hiérarchique des colonnes.
        - popup (tk.Toplevel): Fenêtre popup où les combobox seront placées.

        Returns:
        - str: Chemin complet sélectionné.
        """
        frame = tk.Frame(popup)
        frame.pack()

        # Choix de la colonne principale
        colonne_combo = ttk.Combobox(frame, values=list(feuille.entete.structure.keys()), state="readonly")
        colonne_combo.grid(row=0, column=0, padx=5, pady=5)

        comboboxes = []

        def add_combobox(frame, level, structure, comboboxes):
            combo = ttk.Combobox(frame, state="readonly")
            combo.grid(row=level, column=1, padx=5, pady=2, sticky="w")
            combo["values"] = list(structure.keys())
            comboboxes.append((combo, structure))

            def on_selection(event=None):
                while len(comboboxes) > level + 1:
                    comboboxes[-1][0].destroy()
                    comboboxes.pop()

                selection = combo.get()
                if selection in structure and isinstance(structure[selection], dict) and structure[selection]:
                    add_combobox(frame, level + 1, structure[selection], comboboxes)
                print(f"Selected: {selection}")

            combo.bind("<<ComboboxSelected>>", on_selection)

        def on_colonne_selection(event=None):
            for combo, _ in comboboxes:
                combo.destroy()
            comboboxes.clear()

            selected_col = colonne_combo.get()
            if selected_col in feuille.entete.structure:
                add_combobox(frame, 1, feuille.entete.structure[selected_col], comboboxes)

        colonne_combo.bind("<<ComboboxSelected>>", on_colonne_selection)

        def get_path():
            col1 = colonne_combo.get()
            selection = [combo.get() for combo, _ in comboboxes if combo.get()]
            message = " > ".join([col1] + selection) if col1 else None
            return message
        
        return get_path


# Créer un fichier avec l'en-tête en une seule ligne
    def entete_une_ligne(self):
        """Crée un fichier Excel (.xlsx) avec l'en-tête en une seule ligne."""
        if not self.fichier_path or not self.fichier_path.endswith(".xlsx"):
            messagebox.showerror("Erreur", "Veuillez d'abord sélectionner un fichier .xlsx valide.")
            return
    
        fichier_destination = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Fichiers Excel modernes", "*.xlsx")],
            title="Enregistrer le fichier converti",
            initialdir="sauvegardes/results",
            initialfile=opti_fichier.fichier_du_chemin(self.fichier_path)
        )
        if not fichier_destination:
            return
    
        try:
            fichier = Fichier(self.fichier_path)
            feuille = Feuille(fichier, self.feuille_nom.get(),
                            self.details_structure["data_debut"],
                            self.details_structure["data_fin"],)
            entete = Entete(feuille,self.details_structure["entete_debut"],
                            self.details_structure["entete_fin"],
                            self.details_structure["nb_colonnes_secondaires"],
                            self.details_structure["ligne_unite"],
                            self.dico_entete()
                            )
            feuille.entete = entete
            opti_xlsx.entete_une_ligne(feuille,fichier_destination)
            self.status_label.config(text="✅ creation terminée avec succès", fg="green")
            messagebox.showinfo("Succès", f"Fichier creer avec succès :\n{fichier_destination}")
        except Exception as e:
            self.status_label.config(text="❌ Échec de la creation", fg="red")
            messagebox.showerror("Erreur", f"Erreur lors de la creation : {e}")



# Affichage de l'aperçu du fichier Excel
    def create_excel_preview_frame(self):
        """Crée le cadre pour l'aperçu du fichier Excel."""
        # Créer un LabelFrame
        self.excel_preview_frame = ttkb.LabelFrame(self, text="3. Aperçu du fichier Excel")
        self.excel_preview_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Créer le Treeview avec une colonne pour les numéros de ligne
        self.table = ttk.Treeview(self.excel_preview_frame, show="tree headings", height=15,style="Custom.Treeview")
        self.table.grid(row=0, column=0, sticky="nsew")

        # Scrollbars attachées au LabelFrame
        scroll_y = tk.Scrollbar(self.excel_preview_frame, orient="vertical", command=self.table.yview)
        scroll_y.grid(row=0, column=1, sticky='ns')
        self.table.configure(yscrollcommand=scroll_y.set)

        scroll_x = tk.Scrollbar(self.excel_preview_frame, orient="horizontal", command=self.table.xview)
        scroll_x.grid(row=1, column=0, sticky='ew')
        self.table.configure(xscrollcommand=scroll_x.set)

        # Configurer la grille pour que le tableau prenne l'espace
        self.excel_preview_frame.grid_rowconfigure(0, weight=1)
        self.excel_preview_frame.grid_columnconfigure(0, weight=1)

        # Exemple de colonnes (15 colonnes de données)
        nb_cols = 15
        col_names = [f"Col {i+1}" for i in range(nb_cols)]
        self.table["columns"] = col_names

        self.table.heading("#0", text="Ligne", anchor="center")
        self.table.column("#0", width=50, minwidth=30, anchor="center", stretch=False)
        for name in col_names:
            self.table.heading(name, text=name)
            self.table.column(name, anchor="center", width=120, minwidth=100, stretch=True)

        # Exemple de remplissage avec numéros de ligne et valeurs fictives
        for i in range(50):
            values = [f"Valeur {j+1}" for j in range(nb_cols)]
            # Insérer avec le numéro de ligne (text=) et les valeurs
            self.table.insert("", "end", text=str(i + 1), values=values, tags=("ligne",))


        return self.excel_preview_frame

    def on_treeview_configure(self, event):
        """Ajuste la largeur du tableau pour ne pas dépasser 800 pixels."""
        # Limite la largeur à 800 pixels
        max_width = 800
        if self.table.winfo_width() > max_width:
            self.table.config(width=max_width)

    def update_excel(self):
        """Met à jour le tableau avec les données du fichier Excel sélectionné."""
        try:
            self.table.delete(*self.table.get_children())

        # Lire le fichier Excel
            self.df = pd.read_excel(self.fichier_path, sheet_name=self.feuille_nom.get(), header=None).copy()

            nb_cols = len(self.df.columns)
            col_names = [f"Col {i+1}" for i in range(nb_cols)]

            # Réinitialiser les colonnes
            self.table["columns"] = col_names

            for name in col_names:
                self.table.heading(name, text=name)
                self.table.column(name, anchor="center", width=120, minwidth=100, stretch=True)

            # Remplir le tableau
            for i, row in self.df.head(50).iterrows():
                self.table.insert("", "end", text=str(i), values=list(row))
            self.colorier_lignes_range(
                self.details_structure["entete_debut"],
                self.details_structure["entete_fin"])
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de lire le fichier : {e}")
        self.dico_entete()
        
        
    def afficher_excel(self):
        """Affiche le contenu du fichier Excel dans le tableau."""
        try:
            # Vider les anciennes données
            self.taille_entete_entry.delete(0, tk.END)
            self.taille_entete_entry.insert(0, str(1))

            self.update_excel()
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de lire le fichier : {e}")
        self.dico_entete()

    def append_text(self, new_content, color="black"):
        """Ajoute du texte à la zone de résultats sans le remplacer."""
        if not hasattr(self, "result_text"):
            print("Erreur : 'result_text' n'a pas été initialisé.")
            return
        # Créer le tag uniquement s'il n'existe pas
        if color not in self.result_text.tag_names():
            self.result_text.tag_config(color, foreground=color)
        # Insérer le texte avec le tag de couleur
        self.result_text.insert("end", new_content + "\n", color)
        self.result_text.see("end")

    def colorier_ligne(self, ligne_numero, couleur="#FFFF00"):
        """
        Applique une couleur de fond à la ligne spécifiée.
        :param ligne_numero: le numéro de la ligne (1-based comme dans ton exemple)
        :param couleur: couleur en hexadécimal (par exemple, "#FF0000" pour rouge)
        """
        # Créer un tag avec la couleur si pas encore créé
        tag_name = f"ligne_{ligne_numero}"
        if not hasattr(self, 'tags_configures'):
            self.tags_configures = set()
        if tag_name not in self.tags_configures:
            self.table.tag_configure(tag_name, background=couleur)
            self.tags_configures.add(tag_name)

        # Parcourir tous les items pour trouver celui avec le texte correspondant
        for item in self.table.get_children():
            # Vérifier si le texte (le numéro de ligne) correspond
            if self.table.item(item, "text") == str(ligne_numero):
                # Appliquer le tag pour colorier la ligne
                self.table.item(item, tags=(tag_name,))
                break

    def colorier_lignes_range(self, ligne_debut, ligne_fin, couleur="#FFFF00"):
        """
        Colorie toutes les lignes de ligne_debut à ligne_fin en utilisant la fonction colorier_ligne.
        """
        # S'assurer que ligne_debut est inférieur ou égal à ligne_fin
        if ligne_debut > ligne_fin:
            ligne_debut, ligne_fin = ligne_fin, ligne_debut
        for ligne_numero in range(ligne_debut, ligne_fin + 1):
            self.colorier_ligne(ligne_numero, couleur)

    def enlever_toutes_couleurs(self):
        """
        Enlève la coloration de toutes les lignes.
        """
        for item in self.table.get_children():
            # Récupérer tous les tags
            tags = self.table.item(item, "tags")
            # Filtrer pour enlever tous les tags de couleur
            tags = tuple(tag for tag in tags if not tag.startswith("ligne_"))
            self.table.item(item, tags=tags)

    def enlever_couleur_ligne(self, ligne_numero):
        """
        Enlève la coloration de fond appliquée à la ligne spécifiée.
        :param ligne_numero: le numéro de la ligne (1-based comme dans ton exemple)
        """
        for item in self.table.get_children():
            if self.table.item(item, "text") == str(ligne_numero):
                # Récupérer tous les tags de cette ligne
                tags = self.table.item(item, "tags")
                # Supprimer le tag de coloration spécifique
                tags = tuple(tag for tag in tags if not tag.startswith("ligne_"))
                # Mettre à jour l'item sans ces tags
                self.table.item(item, tags=tags)
                break
