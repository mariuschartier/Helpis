import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from fichier import Fichier
from test_gen import Test_gen
from test_spe import Test_spe
from Feuille import Feuille
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
        # print(self.Parent.winfo_width())
        # self.canvas.config(width=1000)  # Définir une largeur fixe (par exemple 400 pixels)

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

        

    def _bind_mousewheel_to_widget(self, widget):
        widget.bind("<Enter>", lambda e: self._set_active_scroll_target(widget))
        widget.bind("<Leave>", lambda e: self._set_active_scroll_target(None))
    
    def _set_active_scroll_target(self, widget):
        self._active_mouse_scroll_target = widget
    
    def _on_mousewheel(self, event):
        if hasattr(self, "_active_mouse_scroll_target") and self._active_mouse_scroll_target:
            target = self._active_mouse_scroll_target
            if isinstance(target, tk.Text) or isinstance(target, ttk.Treeview):
                target.yview_scroll(int(-1*(event.delta/120)), "units")
        else:
            self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            
            
            
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
    
        # Choix de la taille de l'en-tête
        tk.Label(self.file_frame, text="Taille de l'en-tête :").pack(side="left", padx=(10, 0))
        self.taille_entete_entry = tk.Entry(self.file_frame, width=5)
        self.taille_entete_entry.pack(side="left", padx=5)
        tk.Button(self.file_frame, text="❓ Aide", command=self.ouvrir_aide).pack(side="right", padx=5)
        self.taille_entete_entry.bind("<KeyRelease>", self.on_key_release)
    
        return self.file_frame  # Retourne le cadre créé

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
    
        self._bind_mousewheel_to_widget(self.table)

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
        self._bind_mousewheel_to_widget(self.result_text)
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

        self._bind_mousewheel_to_widget(self.erreur_table)

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

    def afficher_excel(self):
        self.table.delete(*self.table.get_children())
        try:
            self.df = pd.read_excel(self.fichier_path, sheet_name=self.feuille_nom.get(), header=None)
            self.table["columns"] = list(range(len(self.df.columns)))
            self.colonnes_disponibles = list(self.df.iloc[0].dropna().astype(str))
            self.table["show"] = "headings"
            for col in self.table["columns"]:
                self.table.heading(col, text=f"Col {col}")
            for i, row in self.df.head(100).iterrows():
                self.table.insert("", "end", values=list(row))
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de lire la feuille : {e}")

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
        
        # Supprimer les doublons tout en conservant l'ordre
        colonnes_uniques = []
        seen = set()
        for col in self.colonnes_disponibles:
            if col not in seen:
                colonnes_uniques.append(col)
                seen.add(col)
                
                
                
        popup = tk.Toplevel(self)
        popup.title("Ajouter un test spécifique")
    
        tk.Label(popup, text="Nom du test :").grid(row=0, column=0, sticky="w")
        nom_entry = tk.Entry(popup, width=30)
        nom_entry.grid(row=0, column=1)
    
        tk.Label(popup, text="Type de test :").grid(row=1, column=0, sticky="w")
        type_test = ttk.Combobox(popup, values=["val_min", "val_max", "val_entre", "compare_fix", "compare_ratio"], state="readonly")
        type_test.grid(row=1, column=1)
        type_test.set("val_min")
    
        tk.Label(popup, text="Colonne 1 :").grid(row=2, column=0, sticky="w")
        col1_combo = ttk.Combobox(popup, values=colonnes_uniques, state="readonly")
        col1_combo.grid(row=2, column=1)
    


        tk.Label(popup, text="Colonne 2 (si comparaison) :").grid(row=3, column=0, sticky="w")
        col2_combo = ttk.Combobox(popup, values=colonnes_uniques, state="readonly")
        col2_combo.grid(row=3, column=1)

    
        # Champs classiques
        tk.Label(popup, text="Valeur min / différence / ratio :").grid(row=4, column=0, sticky="w")
        val1_entry = tk.Entry(popup)
        val1_entry.grid(row=4, column=1)
    
        tk.Label(popup, text="Valeur max (si besoin) :").grid(row=5, column=0, sticky="w")
        val2_entry = tk.Entry(popup)
        val2_entry.grid(row=5, column=1)
    
        # ✅ Checkboxes pour lire mini/maxi depuis la feuille
        lire_min_var = tk.BooleanVar()
        lire_max_var = tk.BooleanVar()
    
        lire_min_check = tk.Checkbutton(popup, text="Lire la valeur min depuis la feuille (avant-dernière ligne)", variable=lire_min_var)
        lire_min_check.grid(row=6, column=0, columnspan=2, sticky="w")
    
        lire_max_check = tk.Checkbutton(popup, text="Lire la valeur max depuis la feuille (avant-dernière ligne)", variable=lire_max_var)
        lire_max_check.grid(row=7, column=0, columnspan=2, sticky="w")
    
        def ajouter():
            nom = nom_entry.get().strip()
            type_selected = type_test.get()
            col1 = col1_combo.get().strip()
            col2 = col2_combo.get().strip()
    
            try:
                val1 = float(val1_entry.get()) if val1_entry.get() else None
                val2 = float(val2_entry.get()) if val2_entry.get() else None
            except ValueError:
                messagebox.showerror("Erreur", "Valeurs numériques invalides")
                return
    
            if not nom or not col1:
                messagebox.showerror("Erreur", "Nom et colonne principale requis")
                return
    
            test = Test_spe(nom=nom, feuille=None)
    
            # Ajoute aussi les cases à cocher à la suite
            self.tests.append((
                test, type_selected, col1, col2, val1, val2,
                lire_min_var.get(), lire_max_var.get()
            ))
    
            self.test_listbox.insert(tk.END, f"[SPE] {nom} ({type_selected})")
            popup.destroy()
    
        tk.Button(popup, text="Ajouter le test", command=ajouter).grid(row=8, column=1, pady=10)


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

        feuille = Feuille(Fichier(self.fichier_path), self.feuille_nom.get(), taille_entete)
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
                obj, type_test, col1, col2, val1, val2, lire_min, lire_max = test
                obj.feuille = feuille  # mise à jour de la feuille
            
                self.result_text.insert(tk.END, f"--- {obj.nom} ({type_test}) ---\n")
            
                try:
                    min_max =None
                    if lire_min or lire_max:
                        if not( lire_min and lire_max):
                            if lire_max:
                                min_max = "maxi"
                            else:
                                min_max = "mini"
            
                    # ⬇️ Appel normal
                    if type_test == "val_min":
                        message = obj.val_min(val1, col1, min_max)
                    elif type_test == "val_max":
                        message = obj.val_max(val1, col1, min_max)
                    elif type_test == "val_entre":
                        message = obj.val_entre(val1, val2, col1, min_max)
                    elif type_test == "compare_fix":
                        message = obj.compare_col_fix(val1, col1, col2, min_max)
                    elif type_test == "compare_ratio":
                        message = obj.compare_col_ratio(val1, col1, col2, min_max)
            
                    self.result_text.insert(tk.END, str(message) + "\n")
            
                except Exception as e:
                    self.result_text.insert(tk.END, f"Erreur test {obj.nom}: {e}\n")



            self.result_text.insert(tk.END, "\n")
        feuille.error_all_cell_colors()
        for row_idx, ligne in enumerate(feuille.erreurs):
            for col_idx, code in enumerate(ligne):
                if code > 0:
                    self.erreur_table.insert("", "end", values=(row_idx + 1, col_idx + 1, code))


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelTesterApp(root)
    root.mainloop()
