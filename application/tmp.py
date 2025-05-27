from tests import opti_fichier  # ton module de conversion
from tkinter import filedialog, messagebox, ttk
import tkinter as tk
from pathlib import Path
from tests import opti_xlsx 
from tests import opti_separation
import pandas as pd

from structure.Fichier import Fichier
from structure.Feuille import Feuille
import os
from structure.Entete import Entete

class opti_xls(tk.Frame):
    def __init__(self, parent, controller):
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
        
        self.choix_page()
        self.champs_xls()
        self.champs_xlsx()
        self.champs_separation()

        self.status_label = tk.Label(self, text="", bg="#f4f4f4", fg="green")
        self.status_label.pack(pady=5)
        # self.create_excel_preview_frame()
    

    def champs_xls(self):
        frame_action = tk.LabelFrame(self, text="2. action sur le fichier xls", bg="#f4f4f4")
        frame_action.pack(fill="x", padx=10, pady=5)
    
        self.btn_convertir = tk.Button(frame_action, text="Convertir en .xlsx", command=self.controller.bind_button(self.convertir_fichier))
        self.btn_convertir.pack(side="left", padx=5)
        self.btn_convertir.config(state="disabled")

    def champs_xlsx(self):
        frame_action = tk.LabelFrame(self, text="3. action sur le fichier xlsx", bg="#f4f4f4")
        frame_action.pack(fill="x", padx=10, pady=5)

        self.btn_ameliorer = tk.Button(frame_action, text="ameliorer le .xlsx", command=self.controller.bind_button(self.ameliorer_fichier_xlsx))
        self.btn_ameliorer.pack(side="left", padx=5)

        self.btn_moy_jour = tk.Button(frame_action, text="moyenne par jour", command=self.controller.bind_button(self.moyenne_par_jour))
        self.btn_moy_jour.pack(side="left", padx=5)

        self.btn_moy_semaine = tk.Button(frame_action, text="moyenne par semaine", command=self.controller.bind_button(self.moyenne_par_semaine))
        self.btn_moy_semaine.pack(side="left", padx=5)
        self.btn_entete_une_ligne = tk.Button(frame_action, text="Entete en une ligne", command=self.controller.bind_button(self.entete_une_ligne))
        self.btn_entete_une_ligne.pack(side="left", padx=5)

        self.btn_ameliorer.config(state="disabled")
        self.btn_moy_jour.config(state="disabled")
        self.btn_moy_semaine.config(state="disabled")
        self.btn_entete_une_ligne.config(state="disabled")


    def champs_separation(self):
        frame_action = tk.LabelFrame(self, text="4. Création de ficier séparer (xlsx)", bg="#f4f4f4")
        frame_action.pack(fill="x", padx=10, pady=5)
        self.btn_separation = tk.Button(frame_action, text="séparation valeur dans la colonne", command=self.controller.bind_button(self.split_excel_by_column))
        self.btn_separation.pack(side="left", padx=5)

        self.btn_separation.config(state="disabled")



    def choix_page(self):
        self.frame_fichier = tk.LabelFrame(self, text="1. Charger un fichier Excel", bg="#f4f4f4")
        self.frame_fichier.pack(fill="x", padx=10, pady=5)

        self.fichier_entry = tk.Entry(self.frame_fichier, width=80)
        self.fichier_entry.pack(side="left", padx=5, pady=5)

        tk.Button(self.frame_fichier, text="Parcourir", command=self.controller.bind_button(self.choisir_fichier)).pack(side="left", padx=5)
        # Choix de la feuille
        self.feuille_combo = ttk.Combobox(self.frame_fichier, textvariable=self.feuille_nom, state="readonly")
        self.feuille_combo.pack(side="left", padx=5)

        # Choix de la taille de l'en-tête
        self.taille_entete_var = tk.StringVar()
        tk.Label(self.frame_fichier, text="Taille de l'en-tête :").pack(side="left", padx=(10, 0))
        self.taille_entete_entry = tk.Entry(self.frame_fichier, width=5,textvariable=self.taille_entete_var )
        self.taille_entete_var.trace_add("write", self.on_taille_entete_change)

        self.taille_entete_var.set(1)  # Met à jour l'Entry avec 1

        self.taille_entete_entry.pack(side="left", padx=5)
        tk.Button(self.frame_fichier, text="❓ Aide", command=self.ouvrir_aide).pack(side="right", padx=5)
        self.taille_entete_entry.bind("<KeyRelease>", self.on_key_release)

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

        
    def prepare_dossiers(self):
        Path("sauvegardes_tests").mkdir(exist_ok=True)
        Path("results").mkdir(exist_ok=True)
        Path("data").mkdir(exist_ok=True)
        
    def choisir_fichier(self):
        dossier_data = Path("results")
        dossier_data.mkdir(parents=True, exist_ok=True)  # Crée le dossier s’il n’existe pas

        filepath = filedialog.askopenfilename(
            filetypes=[("Fichiers Excel", "*.xls;*.xlsx")],
            initialdir=dossier_data,  # Dossier par défaut
            title="Choisir un fichier"
        )

        if filepath:
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


                self.btn_convertir.config(state="disabled")

            elif is_xls:
                self.btn_convertir.config(state="normal")

                self.btn_ameliorer.config(state="disabled")
                self.btn_moy_jour.config(state="disabled")
                self.btn_moy_semaine.config(state="disabled")
                self.btn_separation.config(state="disabled")
                self.btn_entete_une_ligne.config(state="disabled")

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

    def ouvrir_aide(self):
        aide_popup = tk.Toplevel(self)
        aide_popup.title("Aide - Utilisation")
        aide_popup.geometry("600x400")

        texte = tk.Text(aide_popup, wrap="word", font=("Segoe UI", 10))
        texte.pack(fill="both", expand=True, padx=10, pady=10)

        contenu = (
            "🔍 Bienvenue dans l'application Testeur Excel\n\n"
            "Voici comment utiliser l'application :\n"
            "1️⃣ Cliquez sur 'Parcourir' pour charger un fichier Excel (.xlsx)\n"
            "2️⃣ Choisissez la feuille à analyser dans la liste déroulante\n"
            "3️⃣ Indiquez la taille de l’en-tête (nombre de lignes au début du tableau)\n"
            "4️⃣ Ajoutez un test générique (valeur minimale, maximale ou entre) ou spécifique\n"
            "5️⃣ Cliquez sur 'Exécuter les tests' pour analyser le fichier\n\n"
            "💡 Les erreurs sont colorées dans le fichier Excel et listées dans les résultats\n"
            "📌 Vous pouvez faire défiler l’aperçu et les erreurs avec les barres de défilement\n"
        )
        
        texte.insert(tk.END, contenu)

    def on_key_release(self, event):
        if not self.taille_entete_entry.get().isdigit() and self.taille_entete_entry.get() != "":
            messagebox.showwarning("Validation", "Veuillez entrer un nombre entier.")
            self.taille_entete_entry.delete(0, tk.END)



    def afficher_colonne_popup(self, feuille, event=None):
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

        get_path = self.select_column_path(popup,feuille_obj)

        def valider():
            chemin_1 = get_path()
            if not chemin_1:
                messagebox.showerror("Erreur", "Veuillez sélectionner une colonne cible.")
                return
            self.chemin_selectionne = chemin_1
            popup.destroy()

        btn_ok = tk.Button(popup, text="valider", command=valider)
        btn_ok.pack(pady=10)

        self.wait_window(popup)  # Attend que la fenêtre soit fermée

        # Après fermeture, retourner la valeur sélectionnée
        return self.chemin_selectionne


    def select_column_path(self, popup,feuille :Feuille):
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
    



    def dico_entete(self,feuille=None):
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
   








# Methodes de conversion et d'amélioration

    def convertir_fichier(self):
        if not self.fichier_path or not self.fichier_path.endswith(".xls"):
            messagebox.showerror("Erreur", "Veuillez d'abord sélectionner un fichier .xls valide.")
            return
    
        fichier_destination = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Fichiers Excel modernes", "*.xlsx")],
            title="Enregistrer le fichier converti",
            initialdir="results",
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
    
    def ameliorer_fichier_xlsx(self):
        if not self.fichier_path or not self.fichier_path.endswith(".xlsx"):
            messagebox.showerror("Erreur", "Veuillez d'abord sélectionner un fichier .xlsx valide.")
            return
    
        fichier_destination = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Fichiers Excel modernes", "*.xlsx")],
            title="Enregistrer le fichier converti",
            initialdir="results",
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

    def moyenne_par_jour(self):
        if not self.fichier_path or not self.fichier_path.endswith(".xlsx"):
            messagebox.showerror("Erreur", "Veuillez d'abord sélectionner un fichier .xlsx valide.")
            return
    
        fichier_destination = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Fichiers Excel modernes", "*.xlsx")],
            title="Enregistrer le fichier converti",
            initialdir="results",
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
    
    def moyenne_par_semaine(self):
        if not self.fichier_path or not self.fichier_path.endswith(".xlsx"):
            messagebox.showerror("Erreur", "Veuillez d'abord sélectionner un fichier .xlsx valide.")
            return
    
        fichier_destination = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Fichiers Excel modernes", "*.xlsx")],
            title="Enregistrer le fichier converti",
            initialdir="results",
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

    def split_excel_by_column(self):
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
            initialdir="results",
            initialfile=opti_fichier.fichier_du_chemin(self.fichier_path)
        )
        if not fichier_destination:
            return
    
        try:
            

            opti_separation.split_excel_by_column(feuille,colonne, fichier_destination)
            self.status_label.config(text="✅ creation terminée avec succès", fg="green")
            messagebox.showinfo("Succès", f"Fichier creer avec succès :\n{fichier_destination}")
        except Exception as e:
            self.status_label.config(text="❌ Échec de la conversion", fg="red")
            messagebox.showerror("Erreur", f"Erreur lors de la creation : {e}")


    def entete_une_ligne(self):
        if not self.fichier_path or not self.fichier_path.endswith(".xlsx"):
            messagebox.showerror("Erreur", "Veuillez d'abord sélectionner un fichier .xlsx valide.")
            return
    
        fichier_destination = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Fichiers Excel modernes", "*.xlsx")],
            title="Enregistrer le fichier converti",
            initialdir="results",
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

