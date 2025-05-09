import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from structure.Fichier import Fichier
from tests.Test_gen import Test_gen
from tests.Test_spe import Test_spe
from structure.Feuille import Feuille
from structure.Entete import Entete
import os
import json
from pathlib import Path
import threading



class ExcelTesterApp(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        super().__init__(parent)
        self.Parent = parent
   

        
        self.fichier_path = None
        self.tests = []
        self.df = None
        self.feuille_nom = tk.StringVar()
        self.colonnes_disponibles = []

        self.prepare_dossiers()
        self.details_structure = {
            "entete_debut": 0,
            "entete_fin": 1,
            "data_debut": 2,
            "data_fin": None,
            "nb_colonnes_secondaires": 0,
            "ligne_unite": 1,
            "ignorer_vide": True
        }

        # === Canvas + Scroll principal ===
        self.canvas = tk.Canvas(self, bg="#f4f4f4", width=1000, height=600)
        self.scroll_y = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg="#f4f4f4")

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw",width=980)
        self.canvas.configure(yscrollcommand=self.scroll_y.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scroll_y.pack(side="right", fill="y")
        self.Parent.update_idletasks()
        

        self._active_mouse_scroll_target = None
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)


        # === Création des frames dans la scrollable_frame ===
        self.file_frame = self.create_file_frame()  # Assurez-vous que cela retourne un cadre
        self.test_buttons_frame = self.create_test_buttons_frame()  # Et ainsi de suite...
        self.test_list_frame = self.create_test_list_frame()
        self.excel_preview_frame = self.create_excel_preview_frame()
        self.results_frame = self.create_results_frame()
        self.error_details_frame = self.create_error_details_frame()

        # Maintenant, empilez-les, comme ça :
        self.file_frame.pack(fill='both', expand=True, padx=10, pady=5)
        self.test_buttons_frame.pack(fill='both', expand=True, padx=10, pady=5)
        self.test_list_frame.pack(fill='both', expand=True, padx=10, pady=5)
        self.excel_preview_frame.pack(fill='both', expand=True, padx=10, pady=5)
        self.results_frame.pack(fill='both', expand=True, padx=10, pady=5)
        self.error_details_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        
        self.register_scrollable_widgets()
        
        self._bind_mousewheel_to_widget(self.test_listbox)
        self._bind_mousewheel_to_widget(self.result_text)
        self._bind_mousewheel_to_widget(self.table)
        self._bind_mousewheel_to_widget(self.erreur_table)
        
        




    
    def register_scrollable_widgets(self):
        scrollables = [
            self.test_listbox,
            self.result_text,
            self.table,
            self.erreur_table,
        ]
    
        for widget in scrollables:
                self._bind_mousewheel_to_widget(widget)
                
    def _disable_scroll_on_combo(self, widget):
        widget.bind("<Enter>", lambda e: self.canvas.unbind_all("<MouseWheel>"))
        widget.bind("<Leave>", lambda e: self.canvas.bind_all("<MouseWheel>", self._on_mousewheel))
    

    def _bind_mousewheel_to_widget(self, widget):
        widget.bind("<Enter>", lambda e: self._set_active_scroll_target(widget))
        widget.bind("<Leave>", lambda e: self._set_active_scroll_target(None))
    
    def _set_active_scroll_target(self, widget):
        self._active_mouse_scroll_target = widget
    
    def _on_mousewheel(self, event):
        target = getattr(self, "_active_mouse_scroll_target", None)
    
        if isinstance(target, (tk.Text, tk.Listbox, ttk.Treeview)):
            target.yview_scroll(int(-1 * (event.delta / 120)), "units")
        else:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

            
    
            
    def create_file_frame(self):
        self.file_frame = tk.LabelFrame(self.scrollable_frame, text="1. Charger un fichier Excel", bg="#f4f4f4")
        self.file_frame.pack(fill="both", expand=True, padx=10, pady=5)
    
        self.fichier_entry = tk.Entry(self.file_frame, width=80)
        self.fichier_entry.pack(side="left", padx=5, pady=5)
    
        tk.Button(self.file_frame, text="Parcourir", command=self.controller.bind_button(self.choisir_fichier)).pack(side="left", padx=5)
    
        # Choix de la feuille
        self.feuille_combo = ttk.Combobox(self.file_frame, textvariable=self.feuille_nom, state="readonly")
        self.feuille_combo.pack(side="left", padx=5)
        self.feuille_combo.bind("<<ComboboxSelected>>", lambda e: self.afficher_excel())
        
        tk.Button(self.file_frame, text="detail", command=self.ouvrir_popup_manipulation).pack(side="right", padx=5)

    
        # Choix de la taille de l'en-tête
        self.taille_entete_var = tk.StringVar()
        tk.Label(self.file_frame, text="Taille de l'en-tête :").pack(side="left", padx=(10, 0))
        self.taille_entete_entry = tk.Entry(self.file_frame, width=5,textvariable=self.taille_entete_var )
        self.taille_entete_var.trace_add("write", self.on_taille_entete_change)
        self.taille_entete_entry.pack(side="left", padx=5)
        tk.Button(self.file_frame, text="❓ Aide", command=self.ouvrir_aide).pack(side="right", padx=5)
        self.taille_entete_entry.bind("<KeyRelease>", self.on_key_release)
    
        return self.file_frame  # Retourne le cadre créé
    
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

                # Mettre à jour l'aperçu des colonnes dans l'interface utilisateur
                self.table["columns"] = list(range(len(self.df.columns)))
                for col in self.table["columns"]:
                    self.table.heading(col, text=f"Col {col}")

                # Réinitialiser les données affichées dans le tableau
                self.table.delete(*self.table.get_children())
                for i, row in self.df.head(100).iterrows():
                    self.table.insert("", "end", values=list(row))

            except Exception as e:
                messagebox.showerror("Erreur", f"Impossible de mettre à jour les colonnes : {e}")        

    
    def ouvrir_popup_manipulation(self):
        if self.df is None:            
            messagebox.showerror("Erreur", "Un fichier doit etre selectionné.")
            return
        popup = tk.Toplevel(self)
        popup.title("Paramètres avancés de la feuille")
        popup.configure(bg="#f4f4f4")
    
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
        tk.Checkbutton(frame_cb, text="Ignorer les lignes vides", variable=ignore_lignes_vides, bg="#f4f4f4").pack(side="left")
    
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
            self.taille_entete_entry.delete(0, tk.END)
            self.taille_entete_entry.insert(0, str(taille_entete))
        
            # Optionnel : garder les valeurs pour un usage futur
            valeurs["ignorer_lignes_vides"] = ignore_lignes_vides.get()
            self.details_structure = valeurs
            popup.destroy()

        tk.Button(frame_btns, text="✅ Appliquer", command=appliquer).pack(side="left", padx=10)
        tk.Button(frame_btns, text="❌ Annuler", command=popup.destroy).pack(side="left", padx=10)




    
        
    
    
    def create_test_buttons_frame(self):
        frame_btn_test = tk.Frame(self.scrollable_frame)
        frame_btn_test.pack(fill="both", expand=True, padx=10, pady=5)

        tk.Button(frame_btn_test, text="Ajouter un test générique", command=self.controller.bind_button(self.popup_ajouter_test_gen)).pack(side="left", padx=10)
        tk.Button(frame_btn_test, text="Ajouter un test spécifique", command=self.controller.bind_button(self.popup_ajouter_test_spe)).pack(side="left", padx=10)
        tk.Button(frame_btn_test, text="Exécuter les tests", command=self.controller.bind_button(self.executer_tests)).pack(side="left", padx=10)
        tk.Button(frame_btn_test, text="💾 Sauvegarder les tests", command=self.controller.bind_button(self.sauvegarder_tests)).pack(side="left", padx=10)
        tk.Button(frame_btn_test, text="📂 Importer des tests", command=self.controller.bind_button(self.importer_tests)).pack(side="left", padx=10)
        return frame_btn_test

    def create_test_list_frame(self):
        self.test_list_frame = tk.LabelFrame(self.scrollable_frame, text="2. Liste des tests", bg="#f4f4f4")
        self.test_list_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Détail des tests
        self.test_listbox = tk.Listbox(self.test_list_frame, height=5, selectmode="extended")
        self.test_listbox.pack(side="left", fill="both", expand=True)

        scrollbar_list = tk.Scrollbar(self.test_list_frame, command=self.test_listbox.yview)
        scrollbar_list.pack(side="right", fill="y")
        self.test_listbox.config(yscrollcommand=scrollbar_list.set)
        self.test_listbox.bind("<Double-Button-1>", self.afficher_details_popup)

        # Retirer les tests
        tk.Button(self.test_list_frame, text="Retirer le test sélectionné", command=self.supprimer_test).pack(pady=5)
        
        return self.test_list_frame

    def create_excel_preview_frame(self):
        # Créer un LabelFrame
        self.excel_preview_frame = tk.LabelFrame(self.scrollable_frame, text="3. Aperçu du fichier Excel", bg="#f4f4f4")
        self.excel_preview_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Créer le Treeview directement dans le LabelFrame
        self.table = ttk.Treeview(self.excel_preview_frame, show='headings', height=10)
        self.table.grid(row=0, column=0, sticky='nsew')
        
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
        
        # Configuration des colonnes
        self.table["columns"] = list(range(15))
        for col in range(15):
            self.table.heading(col, text=f"Col {col}")
            self.table.column(col, width=100)
        
        # Exemple de remplissage
        for i in range(50):
            self.table.insert("", "end", values=[f"Série {i}"] + [f"Valeur {j}" for j in range(14)])
    


        return self.excel_preview_frame

    def on_treeview_configure(self, event):
        # Limite la largeur à 800 pixels
        max_width = 800
        if self.table.winfo_width() > max_width:
            self.table.config(width=max_width)

    def create_results_frame(self):
        self.results_frame = tk.LabelFrame(self.scrollable_frame, text="4. Résultats / Erreurs", bg="#f4f4f4")
        self.results_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.result_text = tk.Text(self.results_frame, height=10, wrap="none")
        self.result_text.pack(fill="both", expand=True, padx=10, pady=5)

        # Barres de défilement pour les résultats
        result_scroll_y = tk.Scrollbar(self.results_frame, command=self.result_text.yview)
        result_scroll_y.pack(side="right", fill="y")
        result_scroll_x = tk.Scrollbar(self.results_frame, orient="horizontal", command=self.result_text.xview)
        result_scroll_x.pack(side="bottom", fill="x")

        self.result_text.configure(yscrollcommand=result_scroll_y.set, xscrollcommand=result_scroll_x.set)
        return self.results_frame


    def create_error_details_frame(self):
        self.error_details_frame = tk.LabelFrame(self.scrollable_frame, text="5. Détails des erreurs", bg="#f4f4f4")
        self.error_details_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.erreur_table = ttk.Treeview(self.error_details_frame, columns=("Ligne", "Colonne", "Code"), show="headings")
        self.erreur_table.heading("Ligne", text="Ligne")
        self.erreur_table.heading("Colonne", text="Colonne")
        self.erreur_table.heading("Code", text="Code d'erreur")
        self.erreur_table.pack(side="left", fill="both", expand=True)

        # Barres de défilement pour les détails des erreurs
        err_scroll_y = tk.Scrollbar(self.error_details_frame, orient="vertical", command=self.erreur_table.yview)
        err_scroll_y.pack(side="right", fill="y")
        err_scroll_x = tk.Scrollbar(self.error_details_frame, orient="horizontal", command=self.erreur_table.xview)
        err_scroll_x.pack(fill="x")

        self.erreur_table.configure(yscrollcommand=err_scroll_y.set, xscrollcommand=err_scroll_x.set)
        return self.error_details_frame


        # texte.config(state="disabled")
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
   
      
    def prepare_dossiers(self):
        Path("sauvegardes_tests").mkdir(exist_ok=True)
        Path("results").mkdir(exist_ok=True)
        Path("data").mkdir(exist_ok=True)
            
    def choisir_fichier(self):
        filepath = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx")],
            initialdir="results",  # Dossier de fichiers à convertir
            title="Choisir un fichier Excel"
        )
        if filepath:
            self.fichier_path = filepath
            self.fichier_entry.delete(0, tk.END)
            self.fichier_entry.insert(0, filepath)
            try:
                xls = pd.ExcelFile(filepath)
                self.feuille_combo['values'] = xls.sheet_names
                self.feuille_combo.set(xls.sheet_names[0])
                self.afficher_excel()
            except Exception as e:
                messagebox.showerror("Erreur", f"Impossible de lire les feuilles du fichier : {e}")


    def validate_integer_input(self, P):
        return P == "" or P.isdigit()

    def on_key_release(self, event):
        if not self.taille_entete_entry.get().isdigit() and self.taille_entete_entry.get() != "":
            messagebox.showwarning("Validation", "Veuillez entrer un nombre entier.")
            self.taille_entete_entry.delete(0, tk.END)

    def dico_entete(self):
        self.dico_structure = {}
        ligne_entete_debut = self.details_structure.get("entete_debut", 0)
        ligne_entete_fin = self.details_structure.get("entete_fin", 1)

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
            # messagebox.showerror("Erreur", f"Impossible de construire le dictionnaire d'en-tête : {e}")
            return {}

        

    def afficher_excel(self):
        self.table.delete(*self.table.get_children())
        try:
            self.df = pd.read_excel(self.fichier_path, sheet_name=self.feuille_nom.get(), header=None)

            self.table["columns"] = list(range(len(self.df.columns)))
    
            # 🔍 Récupère les colonnes depuis la bonne ligne d'en-tête
            ligne_entete_debut = self.details_structure.get("entete_debut", 0)
            self.colonnes_disponibles = list(self.df.iloc[ligne_entete_debut].dropna().astype(str))
    
            self.table["show"] = "headings"
            for col in self.table["columns"]:
                self.table.heading(col, text=f"Col {col}")
    
            for i, row in self.df.head(100).iterrows():
                self.table.insert("", "end", values=list(row))
    
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de lire la feuille : {e}")
        self.dico_entete()
        


    def supprimer_test(self):
        selection = self.test_listbox.curselection()
        if not selection:
            return
    
        # Supprimer dans l'ordre inverse pour éviter les décalages d'index
        for index in reversed(selection):
            self.test_listbox.delete(index)
            del self.tests[index]
    
        self.result_text.insert(tk.END, f"{len(selection)} test(s) supprimé(s).\n")
    
                

    
    def sauvegarder_tests(self):
        Path("sauvegardes_tests").mkdir(exist_ok=True)
        from tkinter import filedialog
        chemin = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
            initialdir="sauvegardes_tests",  # Répertoire par défaut
            title="Sauvegarder les tests"
        )
        if not chemin:
            return
    
        export = []
        for test in self.tests:
            if isinstance(test[0], Test_gen):
                obj, type_test, val_min, val_max = test
                export.append({
                    "type": "gen",
                    "nom": obj.nom,
                    "critere": obj.critere,
                    "test_type": type_test,
                    "val_min": val_min,
                    "val_max": val_max
                })
            elif isinstance(test[0], Test_spe):
                obj, test_type, col1, col2, val1, val2, *options = test
                lire_min = options[0] if len(options) > 0 else False
                lire_max = options[1] if len(options) > 1 else False
                export.append({
                    "type": "spe",
                    "nom": obj.nom,
                    "test_type": test_type,
                    "col1": col1,
                    "col2": col2,
                    "val1": val1,
                    "val2": val2,
                    "lire_min": lire_min,
                    "lire_max": lire_max
                })

    
        with open(chemin, "w", encoding="utf-8") as f:
            json.dump(export, f, ensure_ascii=False, indent=2)

            
            
    def importer_tests(self):
        chemin = filedialog.askopenfilename(
            filetypes=[("JSON", "*.json")],
            initialdir="sauvegardes_tests",  # Répertoire par défaut
            title="Importer un fichier de tests"
        )
        if not chemin:
            return
        
    
        try:
            with open(chemin, "r", encoding="utf-8") as f:
                data = json.load(f)
    
            for test in data:
                if test["type"] == "gen":
                    obj = Test_gen(nom=test["nom"], critere=test["critere"])
                    self.tests.append((obj, test["test_type"], test.get("val_min"), test.get("val_max")))
                    self.test_listbox.insert(tk.END, f"[GEN] {test['nom']} ({test['test_type']})")
                elif test["type"] == "spe":
                    feuille = None
                    obj = Test_spe(nom=test["nom"], feuille=feuille)
                    self.tests.append((
                        obj,
                        test["test_type"],
                        test["col1"],
                        test["col2"],
                        test.get("val1"),
                        test.get("val2"),
                        test.get("lire_min", False),
                        test.get("lire_max", False)
                    ))
                    self.test_listbox.insert(tk.END, f"[SPE] {test['nom']} ({test['test_type']})")

        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de charger les tests : {e}")

    
    def afficher_details_popup(self, event):
        selection = self.test_listbox.curselection()
        if not selection:
            return
    
        index = selection[0]
        test_info = self.tests[index]
    
        popup = tk.Toplevel(self)
        popup.title("Détails du test")
        popup.geometry("350x250")
    
        if isinstance(test_info[0], Test_spe):
            # 🛠️ Support des nouveaux champs
            test_obj, test_type, col1, col2, val1, val2, lire_min, lire_max = test_info
            type_label = f"Type : {test_type}"
    
            champs = {
                "val_min": [("Colonne", col1), ("Valeur Min", val1), ("Lire min depuis fichier", lire_min)],
                "val_max": [("Colonne", col1), ("Valeur Max", val1), ("Lire max depuis fichier", lire_max)],
                "val_entre": [("Colonne", col1), ("Valeur Min", val1), ("Lire min", lire_min), ("Valeur Max", val2), ("Lire max", lire_max)],
                "compare_fix": [("Colonne 1", col1), ("Colonne 2", col2), ("Différence max", val1)],
                "compare_ratio": [("Colonne 1", col1), ("Colonne 2", col2), ("Ratio autorisé", val1)],
            }
    
        elif isinstance(test_info[0], Test_gen):
            test_obj, test_type, val_min, val_max = test_info
            type_label = f"Type : {test_type}"
            champs = {
                "val_min": [("Critères", ", ".join(test_obj.critere)), ("Valeur Min", val_min)],
                "val_max": [("Critères", ", ".join(test_obj.critere)), ("Valeur Max", val_max)],
                "val_entre": [("Critères", ", ".join(test_obj.critere)), ("Valeur Min", val_min), ("Valeur Max", val_max)],
            }
    
        tk.Label(popup, text=f"Nom : {test_obj.nom}", font=("Segoe UI", 10, "bold")).pack(pady=5)
        tk.Label(popup, text=type_label).pack(pady=5)
    
        for champ, valeur in champs.get(test_type, []):
            tk.Label(popup, text=f"{champ} : {valeur}").pack(anchor="w", padx=20)



    def popup_ajouter_test_gen(self):


        # Supprimer les doublons tout en conservant l'ordre
        colonnes_uniques = []
        seen = set()
        for col in self.colonnes_disponibles:
            if col not in seen:
                colonnes_uniques.append(col)
                seen.add(col)

        tk.Label(popup, text="Colonne à tester :").grid(row=2, column=0, sticky="w")
        col_test_combobox = ttk.Combobox(popup, values=colonnes_uniques, state="readonly")
        col_test_combobox.grid(row=2, column=1)




        popup = tk.Toplevel(self)
        popup.title("Ajouter un test générique")

        tk.Label(popup, text="Nom du test :").grid(row=0, column=0, sticky="w")
        nom_entry = tk.Entry(popup, width=30)
        nom_entry.grid(row=0, column=1)

        tk.Label(popup, text="Critères (séparés par des virgules) :").grid(row=1, column=0, sticky="w")
        critere_entry = tk.Entry(popup, width=40)
        critere_entry.grid(row=1, column=1)

        tk.Label(popup, text="Type de test :").grid(row=2, column=0, sticky="w")
        type_test = ttk.Combobox(popup, values=["val_min", "val_max", "val_entre"], state="readonly")
        type_test.grid(row=2, column=1)
        type_test.set("val_entre")

        tk.Label(popup, text="Valeur minimum :").grid(row=3, column=0, sticky="w")
        val_min_entry = tk.Entry(popup)
        val_min_entry.grid(row=3, column=1)

        tk.Label(popup, text="Valeur maximum :").grid(row=4, column=0, sticky="w")
        val_max_entry = tk.Entry(popup)
        val_max_entry.grid(row=4, column=1)

        def ajouter():
            nom = nom_entry.get().strip()
            criteres = [c.strip() for c in critere_entry.get().split(",") if c.strip()]
            type_selected = type_test.get()
            try:
                val_min = float(val_min_entry.get()) if val_min_entry.get() else None
                val_max = float(val_max_entry.get()) if val_max_entry.get() else None
            except ValueError:
                messagebox.showerror("Erreur", "Valeurs numériques invalides")
                return

            if not nom or not criteres:
                messagebox.showerror("Erreur", "Nom et critères requis")
                return

            test = Test_gen(nom=nom, critere=criteres)
            self.tests.append((test, type_selected, val_min, val_max))
            self.test_listbox.insert(tk.END, f"[GEN] {nom} ({type_selected})")
            popup.destroy()

        tk.Button(popup, text="Ajouter le test", command=ajouter).grid(row=5, column=1, pady=10)

    def popup_ajouter_test_spe(self):
        try:
            dico = self.dico_entete()  # Assure que self.dico_structure est construit
        except Exception as e:
            messagebox.showerror("Erreur", "Fichier et taille d'entete requis.")
            return
        if dico == {}:
            return

        popup = tk.Toplevel(self)
        popup.title("Ajouter un test spécifique")

        # Nom du test
        tk.Label(popup, text="Nom du test :").grid(row=0, column=0, sticky="w")
        nom_entry = tk.Entry(popup, width=30)
        nom_entry.grid(row=0, column=1)

        # Type de test
        tk.Label(popup, text="Type de test :").grid(row=1, column=0, sticky="w")
        type_test = ttk.Combobox(popup, values=["val_min", "val_max", "val_entre", "compare_fix", "compare_ratio"], state="readonly")
        type_test.grid(row=1, column=1)
        type_test.set("val_min")

        # Choix de la première colonne cible
        label_col1 = tk.Label(popup, text="Colonne cible 1 :")
        label_col1.grid(row=2, column=0, sticky="w")
        colonne_cible_1_combo = ttk.Combobox(popup, state="readonly")
        colonne_cible_1_combo.grid(row=2, column=1)
        colonne_cible_1_combo["values"] = list(self.dico_structure.keys())


        # Cadres pour les sous-catégories des deux colonnes
        label_combo1 = tk.Label(popup, text="Sous-catégories 1 :")
        label_combo1.grid(row=3, column=0, sticky="w")
        frame_comboboxes_1 = tk.Frame(popup)
        frame_comboboxes_1.grid(row=4, column=0, columnspan=2, sticky="w")

        # Choix de la deuxième colonne cible
        label_col2 = tk.Label(popup, text="Colonne cible 2 :")
        colonne_cible_2_combo = ttk.Combobox(popup, state="readonly")
        colonne_cible_2_combo["values"] = list(self.dico_structure.keys())


        label_combo2 =tk.Label(popup, text="Sous-catégories 2 :")
        frame_comboboxes_2 = tk.Frame(popup)

        comboboxes_1 = []
        comboboxes_2 = []

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

            combo.bind("<<ComboboxSelected>>", on_selection)

        def on_colonne_selection(col_combo, frame, comboboxes):
            for combo, _ in comboboxes:
                combo.destroy()
            comboboxes.clear()

            selected_col = col_combo.get()
            if selected_col in self.dico_structure:
                add_combobox(frame, 0, self.dico_structure[selected_col], comboboxes)

        colonne_cible_1_combo.bind("<<ComboboxSelected>>", lambda e: on_colonne_selection(colonne_cible_1_combo, frame_comboboxes_1, comboboxes_1))
        colonne_cible_2_combo.bind("<<ComboboxSelected>>", lambda e: on_colonne_selection(colonne_cible_2_combo, frame_comboboxes_2, comboboxes_2))

        # Champs dynamiques selon le type de test
        ligne = 5 # derniere ligne  

        label_val_min = tk.Label(popup, text="Valeur minimale :")
        val_min_entry = tk.Entry(popup)

        label_val_max = tk.Label(popup, text="Valeur maximale :")
        val_max_entry = tk.Entry(popup)

        label_diff = tk.Label(popup, text="Différence attendue :")
        diff_entry = tk.Entry(popup)

        label_ratio = tk.Label(popup, text="Ratio attendu :")
        ratio_entry = tk.Entry(popup)

        for widget in [label_val_min, val_min_entry, label_val_max, val_max_entry, label_diff, diff_entry, label_ratio, ratio_entry, label_col2, colonne_cible_2_combo, label_combo2, frame_comboboxes_2]:
            widget.grid_forget()

        def afficher_champs_selon_type(event=None):
            for widget in [label_val_min, val_min_entry, label_val_max, val_max_entry, label_diff, diff_entry, label_ratio, ratio_entry,  label_col2, colonne_cible_2_combo, label_combo2, frame_comboboxes_2]:
                widget.grid_forget()
            ligne = 5 # derniere ligne  

            t = type_test.get()
            ligne_i = ligne

            if t == "val_min":
                label_val_min.grid(row=ligne_i, column=0, sticky="w")
                val_min_entry.grid(row=ligne_i, column=1)
            elif t == "val_max":
                label_val_max.grid(row=ligne_i, column=0, sticky="w")
                val_max_entry.grid(row=ligne_i, column=1)
            elif t == "val_entre":
                label_val_min.grid(row=ligne_i, column=0, sticky="w")
                val_min_entry.grid(row=ligne_i, column=1)
                ligne_i += 1
                label_val_max.grid(row=ligne_i, column=0, sticky="w")
                val_max_entry.grid(row=ligne_i, column=1)
            elif t =="compare_fix"or t =="compare_ratio":
                label_col2.grid(row=ligne_i, column=0, sticky="w")
                colonne_cible_2_combo.grid(row=ligne_i, column=1)
                ligne_i += 1
                label_combo2.grid(row=ligne_i, column=0, sticky="w")
                frame_comboboxes_2.grid(row=ligne_i+1, column=0, columnspan=2, sticky="w")
                ligne_i +=2
                label_val_min.grid(row=ligne_i, column=0, sticky="w")
                val_min_entry.grid(row=ligne_i, column=1)
                ligne +=1
                if t == "compare_fix":
                    label_diff.grid(row=ligne_i, column=0, sticky="w")
                    diff_entry.grid(row=ligne_i, column=1)
                    ligne +=1

                else:
                    label_ratio.grid(row=ligne_i, column=0, sticky="w")
                    ratio_entry.grid(row=ligne_i, column=1)
                    ligne +=1



        type_test.bind("<<ComboboxSelected>>", afficher_champs_selon_type)
        afficher_champs_selon_type()

        def ajouter():
            nom = nom_entry.get().strip()
            type_selected = type_test.get()

            col1 = colonne_cible_1_combo.get()
            col2 = colonne_cible_2_combo.get() if colonne_cible_2_combo.winfo_ismapped() else None

            # Collecte des sous-catégories sélectionnées
            selection_1 = [combo.get() for combo, _ in comboboxes_1 if combo.get()]
            selection_2 = [combo.get() for combo, _ in comboboxes_2 if combo.get()]

            # Chemins complets des colonnes
            chemin_1 = " > ".join([col1] + selection_1) if col1 else None
            chemin_2 = " > ".join([col2] + selection_2) if col2 else None

            try:
                val1 = float(val_min_entry.get()) if val_min_entry.get() else None
                val2 = float(val_max_entry.get()) if val_max_entry.get() else None
            except ValueError:
                messagebox.showerror("Erreur", "Valeurs numériques invalides")
                return

            if not nom or not chemin_1:
                messagebox.showerror("Erreur", "Nom et Colonne cible 1 requis")
                return

            test = Test_spe(nom=nom, feuille=None)
            self.tests.append((test, type_selected, chemin_1, chemin_2, val1, val2))
            self.test_listbox.insert(tk.END, f"[SPE] {nom} ({type_selected})")
            popup.destroy()

        tk.Button(popup, text="Ajouter le test", command=ajouter).grid(row=ligne + 6, column=1, pady=10)






    def executer_tests(self):
        taille_entete_str = self.taille_entete_entry.get()
        if size_str := taille_entete_str.strip():
            try:
                taille_entete = int(size_str)
            except ValueError:
                messagebox.showerror("Erreur", "Veuillez entrer une valeur entière valide.")
                return
        else:
            messagebox.showwarning("Attention", "Veuillez entrer une taille d'en-tête.")
            return

        if not self.fichier_path or not self.feuille_nom.get():
            messagebox.showerror("Erreur", "Aucun fichier ou feuille sélectionné.")
            return

        fichier = Fichier(self.fichier_path)
        feuille = Feuille(fichier, self.feuille_nom.get(),
                          self.details_structure["data_debut"],
                          self.details_structure["data_fin"],)
        entete = Entete(feuille,self.details_structure["entete_debut"],
                        self.details_structure["entete_fin"],
                        self.details_structure["nb_colonnes_secondaires"],
                        self.details_structure["ligne_unite"],
                        self.dico_structure
                        )
        feuille.entete = entete
        # print(feuille.entete.structure)
        feuille.clear_all_cell_colors()


        self.result_text.delete("1.0", tk.END)
        for item in self.erreur_table.get_children():
            self.erreur_table.delete(item)

        for test in self.tests:
            message = ""
            if isinstance(test[0], Test_gen):
                obj, type_test, val_min, val_max = test
                self.result_text.insert(tk.END, f"--- {obj.nom} ({type_test}) fini---\n")
                try:
                    if type_test == "val_min":
                        message = obj.val_min(feuille, val_min)
                    elif type_test == "val_max":
                        message = obj.val_max(feuille, val_max)
                    elif type_test == "val_entre":
                        message = obj.val_entre(feuille, val_min, val_max)
                    
                    self.result_text.insert(tk.END, str(message) + "\n")

                    

                except Exception as e:
                    self.result_text.insert(tk.END, f"Erreur test {obj.nom}: {e}")


            elif isinstance(test[0], Test_spe):
                # ⬇️ décomposition étendue avec les nouvelles cases à cocher
                obj, type_test, col1, col2, val1, val2 = test
                obj.feuille = feuille  # mise à jour de la feuille
            
                self.result_text.insert(tk.END, f"--- {obj.nom} ({type_test}) ---\n")
            
                try:
            
                    # ⬇️ Appel normal
                    if type_test == "val_min":
                        message = obj.val_min(val1, col1)
                    elif type_test == "val_max":
                        message = obj.val_max(val1, col1)
                    elif type_test == "val_entre":
                        message = obj.val_entre(val1, val2, col1)
                    elif type_test == "compare_fix":
                        message = obj.compare_col_fix(val1, col1, col2)
                    elif type_test == "compare_ratio":
                        message = obj.compare_col_ratio(val1, col1, col2)
            
                    self.result_text.insert(tk.END, str(message) + "\n")
            
                except Exception as e:
                    self.result_text.insert(tk.END, f"Erreur test {obj.nom}: {e}\n")



            self.result_text.insert(tk.END, "\n")
        feuille.error_all_cell_colors()
        for row_idx, ligne in enumerate(feuille.erreurs):
            for col_idx, code in enumerate(ligne):
                if code > 0:
                    self.erreur_table.insert("", "end", values=(row_idx + 1, col_idx + 1, code))

