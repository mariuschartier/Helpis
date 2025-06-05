from back.manipulation import opti_fichier  # ton module de conversion
from tkinter import filedialog, messagebox, ttk
import tkinter as tk
import ttkbootstrap as ttkb
from ttkbootstrap.constants import *
from pathlib import Path
from back.manipulation import opti_xlsx 
from back.manipulation import opti_separation
import pandas as pd

from structure.Fichier import Fichier
from structure.Feuille import Feuille
from structure.Selection_col import Selection_col

import os
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
        # self.create_excel_preview_frame()
    
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
        self.feuille_combo.bind("<<ComboboxSelected>>", lambda e: self.afficher_excel())
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
        except Exception as e:
            print(f"erreur lors de la lecture du fichier: {e}")

    def activation_bouton(self,filepath):
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

        else:
            self.btn_convertir.config(state="disabled")
            self.btn_ameliorer.config(state="disabled")
            self.btn_moy_jour.config(state="disabled")
            self.btn_moy_semaine.config(state="disabled")
            self.btn_separation.config(state="disabled")
            self.btn_entete_une_ligne.config(state="disabled")


        if not is_xls:
            try:
                # Sélectionner le moteur approprié
                engine = 'openpyxl' if is_xlsx else None# 'xlrd'
                try:
                    xls = pd.ExcelFile(filepath, engine=engine)
                except ValueError as e:
                    # Si xlrd échoue pour un .xls spécial, essayer avec openpyxl
                    if not is_xlsx:
                        engine = 'openpyxl'
                        try:
                            xls = pd.ExcelFile(filepath, engine=engine)
                        except Exception as e_openpyxl:
                            raise Exception(f"Erreur avec openpyxl : {e_openpyxl}")
                    else:
                        raise Exception(f"Erreur avec xlrd : {e}")

                feuilles = xls.sheet_names
                self.feuille_combo['values'] = feuilles

                if feuilles:
                    self.feuille_combo['values'] = feuilles
                    self.feuille_nom.set(feuilles[0])  # Met à jour la variable et le combobox

                # Informer de l'extension utilisée
                print("Le fichier est au format .xlsx" if is_xlsx else "Le fichier est au format .xls")

            except Exception as e:
                messagebox.showerror("Erreur", f"Impossible de lire les feuilles du fichier :\n{e}")

    def desactivation_bouton(self):
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
        Met à jour la fin de l'en-tête et reconstruit les colonnes disponibles
        et le dictionnaire d'en-tête en fonction de la nouvelle taille.
        """
        # Mettre à jour la fin de l'en-tête
        self.details_structure["entete_fin"] = (
            int(self.taille_entete_entry.get()) + self.details_structure["entete_debut"] - 1
            if self.taille_entete_entry.get().isdigit()
            else 0
        )
        # Vérifier si un fichier est chargé
        if self.df is not None:
            try:
                # Reconstruire les colonnes disponibles
                ligne_entete_debut = self.details_structure.get("entete_debut", 0)
                self.colonnes_disponibles = list(
                    self.df.iloc[ligne_entete_debut].dropna().astype(str)
                )
                # Reconstruire le dictionnaire d'en-tête
                self.dico_entete()
            except Exception as e:
                messagebox.showerror("Erreur", f"Impossible de mettre à jour les colonnes : {e}")        


    # Ouvrir le popup de manipulation de l'entete detaillée
    def ouvrir_popup_manipulation(self):
        """Ouvre un popup pour configurer les paramètres avancés de la feuille."""
        if self.df is None:            
            messagebox.showerror("Erreur", "Un fichier doit etre selectionné.")
            return
        popup = tk.Toplevel(self)
        popup.title("Paramètres avancés de la feuille")
        popup.configure(bg="#f4f4f4")
        popup.grab_set()

        tk.Label(popup, text="Paramètres de lecture du fichier", font=("Segoe UI", 11, "bold"), bg="#f4f4f4").pack(pady=10)
    
        champs = [
            ("Début de l'en-tête :", "entete_debut"),
            ("Fin de l'en-tête :", "entete_fin"),
            ("Début des données :", "data_debut"),
            ("Fin des données :", "data_fin"),
            ("Colonnes secondaires :", "nb_colonnes_secondaires"),
            ("Ligne des unités :", "ligne_unite"),  # 🆕 Champ ajouté
        ]
    
        entries = {}
        valeurs_par_defaut = self.details_structure if hasattr(self, "details_structure") else {}
    
        for label, key in champs:
            frame = tk.Frame(popup, bg="#f4f4f4")
            frame.pack(fill="x", padx=10, pady=2)
            tk.Label(frame, text=label, width=25, anchor="w", bg="#f4f4f4").pack(side="left")
        
            vcmd = (self.register(lambda val: val.isdigit() or val == ""), '%P')
            entry = tk.Entry(frame, validate="key", validatecommand=vcmd)
            entry.pack(side="left", fill="x", expand=True)
            valeur_defaut = valeurs_par_defaut.get(key, "")
            if key == "data_fin":
                try:
                    valeur_defaut = str(self.df.shape[0])  # Nombre de lignes du DataFrame
                except AttributeError:
                    messagebox.showwarning("Attention", "La feuille de données n'existe pas. La valeur de 'Fin des données' ne peut pas être déterminée.")
                    valeur_defaut = ""
            if key == "data_fin":
                try:
                    entry.insert(0, str(self.df.shape[0]))
                except AttributeError:
                    entry.insert(0, "")
            else:
                entry.insert(0, str(valeur_defaut))  # Initialise avec la valeur par défaut si disponible
            
            entries[key] = entry

        # ✅ Check : ignorer lignes vides (coché par défaut)
        ignore_lignes_vides = tk.BooleanVar(value=True)
        frame_cb = tk.Frame(popup, bg="#f4f4f4")
        frame_cb.pack(padx=10, pady=5, anchor="w")
        tk.Checkbutton(popup, text="Ignorer les lignes vides", variable=ignore_lignes_vides, bg="#f4f4f4").pack(side="left")
        def reset_valeur():
            """Réinitialise les valeurs des champs à leurs valeurs par défaut."""
            for key, entry in entries.items():
                valeur_defaut = valeurs_par_defaut.get(key, "")
                if key == "data_fin":
                    try:
                        entry.delete(0, tk.END)
                        entry.insert(0, str(self.df.shape[0]))
                    except AttributeError:
                        entry.delete(0, tk.END)
                        entry.insert(0, "")
                else:
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
    
        def appliquer():
            try:
                valeurs = {k: int(e.get()) for k, e in entries.items()}
            except ValueError:
                messagebox.showerror("Erreur", "Tous les champs doivent être remplis avec des entiers valides.")
                return
        
            # Calcul automatique de la taille d’en-tête
            taille_entete = valeurs["entete_fin"] - valeurs["entete_debut"] + 1
            if taille_entete <= 0:
                messagebox.showerror("Erreur", "L'entête doit contenir au moins une ligne.")
                return
        
            # Vérification des contraintes
            if valeurs["entete_fin"] >= valeurs["data_debut"]:
                messagebox.showerror("Erreur", "La fin de l'entête doit être avant le début des données.")
                return
        
            if valeurs["nb_colonnes_secondaires"] >= taille_entete:
                messagebox.showerror("Erreur", "Le nombre de colonnes secondaires doit être inférieur à la taille de l'entête.")
                return
        
            if not (valeurs["entete_debut"] <= valeurs["ligne_unite"] <= valeurs["entete_fin"]):
                messagebox.showerror("Erreur", "La ligne d'unité doit être comprise dans l'entête.")
                return
        
            # Appliquer les valeurs
            
        
            # Optionnel : garder les valeurs pour un usage futur
            valeurs["ignorer_lignes_vides"] = ignore_lignes_vides.get()
            self.details_structure = valeurs

            self.taille_entete_entry.delete(0, tk.END)
            self.taille_entete_entry.insert(0, str(taille_entete))
            popup.destroy()

        ttkb.Button(frame_btns, text="✅ Appliquer", command=appliquer).pack(side="left", padx=10)
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
        """Crée les dossiers nécessaires pour l'application."""

        Path("sauvegardes/sauvegardes_tests").mkdir(exist_ok=True)
        Path("sauvegardes/results").mkdir(exist_ok=True)
        Path("sauvegardes/data").mkdir(exist_ok=True)

# Validation de la taille de l'en-tête =========================================================================================================
    def on_key_release_int(self, event):
        """Valide l'entrée de la taille de l'en-tête pour s'assurer qu'elle est un entier positif."""
        if not self.taille_entete_entry.get().isdigit() and self.taille_entete_entry.get() != "":
            messagebox.showwarning("Validation", "Veuillez entrer un nombre entier.")
            self.taille_entete_entry.delete(0, tk.END)


# Construction du dictionnaire d'en-tête =========================================================================================================
    def dico_entete(self,feuille=None):
        """Construit un dictionnaire représentant la structure de l'en-tête du fichier Excel."""

        self.dico_structure = {}
        ligne_entete_debut = self.details_structure.get("entete_debut", 0)
        ligne_entete_fin = self.details_structure.get("entete_fin", 1)

        try:
            for col_idx in range(len(feuille.df.columns)):
                current_level = self.dico_structure

                for row_idx in range(ligne_entete_debut, ligne_entete_fin + 1):
                    cell_value = feuille.df.iloc[row_idx, col_idx]
                    if pd.isna(cell_value):
                        continue
                    cell_value = str(cell_value)

                    if cell_value not in current_level:
                        current_level[cell_value] = {}

                    current_level = current_level[cell_value]

            return self.dico_structure

        except Exception as e:
            messagebox.showerror("Erreur", "Fichier et taille d'entete requis.")
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
    
        try:
            f1 = Fichier(self.fichier_path)
            f1_1 = Feuille(f1,self.feuille_nom.get())
            opti_xlsx.moyenne_par_jour(f1_1,fichier_destination)
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
    
        try:
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
                        self.dico_entete(feuille)
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
                            self.dico_entete(feuille)
                            )
            feuille.entete = entete
            opti_xlsx.entete_une_ligne(feuille,fichier_destination)
            self.status_label.config(text="✅ creation terminée avec succès", fg="green")
            messagebox.showinfo("Succès", f"Fichier creer avec succès :\n{fichier_destination}")
        except Exception as e:
            self.status_label.config(text="❌ Échec de la creation", fg="red")
            messagebox.showerror("Erreur", f"Erreur lors de la creation : {e}")

