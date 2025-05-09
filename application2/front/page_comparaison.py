from tkinter import filedialog, messagebox, ttk,messagebox
import tkinter as tk
import pandas as pd

from tests.ComparateurFichiers import ComparateurFichiers
from tests.fonctions import to_int

from structure.Feuille import Feuille
from structure.Fichier import Fichier


class ComparePage(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg="#f4f4f4")
        self.controller = controller
        self.comparateur = ComparateurFichiers()
        self.feuille_nom = tk.StringVar()
        self.fichier_path = None
        self.df = None

        self.create_file_frame()
        self.block_liste_feuille()

        self.create_excel_preview_frame()
        self.create_test_selector()
        self.create_result_box()


    def create_file_frame(self):
        self.file_frame = tk.LabelFrame(self, text="1. Charger un fichier Excel", bg="#f4f4f4")
        self.file_frame.pack(fill="x", padx=10, pady=5)

        self.fichier_entry = tk.Entry(self.file_frame, width=80)
        self.fichier_entry.pack(side="left", padx=5, pady=5)

        tk.Button(self.file_frame, text="Parcourir", command=self.choisir_fichier).pack(side="left", padx=5)

        self.feuille_combo = ttk.Combobox(self.file_frame, textvariable=self.feuille_nom, state="readonly")
        self.feuille_combo.pack(side="left", padx=5)
        self.feuille_combo.bind("<<ComboboxSelected>>", lambda e: self.afficher_excel())

        tk.Label(self.file_frame, text="Taille de l'en-tête :").pack(side="left", padx=(10, 0))
        self.taille_entete_entry = tk.Entry(self.file_frame, width=5)
        self.taille_entete_entry.pack(side="left", padx=5)

        tk.Button(self.file_frame, text="Ajouter au comparateur", command=self.ajouter_feuille).pack(side="left", padx=10)

    def choisir_fichier(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not path:
            return
        self.fichier_path = path
        self.fichier_entry.delete(0, tk.END)
        self.fichier_entry.insert(0, path)

        try:
            xls = pd.ExcelFile(path)
            self.feuille_combo["values"] = xls.sheet_names
            self.feuille_combo.set(xls.sheet_names[0])
            self.afficher_excel()
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur de lecture du fichier : {e}")
                
    def block_liste_feuille(self):
        # Créer un cadre principal pour la liste et le bouton
        liste_frame = tk.LabelFrame(self, text="Fichiers ajoutés", bg="#f4f4f4")
        liste_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Cadre pour organiser la Listbox et le bouton
        list_bouton_frame = tk.Frame(liste_frame)
        list_bouton_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # La Listbox
        self.liste_feuilles = tk.Listbox(list_bouton_frame, height=5, selectmode="extended")
        self.liste_feuilles.pack(side="left", fill="both", expand=True)

        # Le bouton de suppression à droite
        tk.Button(list_bouton_frame, text="Retirer le test sélectionné", command=self.supprimer_test).pack(side="right", pady=5)


    def supprimer_test(self):
        selection = self.liste_feuilles.curselection()
        if not selection:
            return
    
        # Supprimer dans l'ordre inverse pour éviter les décalages d'index
        for index in reversed(selection):
            self.liste_feuilles.delete(index)
    
        self.result_text.insert(tk.END, f"{len(selection)} test(s) supprimé(s).\n")

    def afficher_excel(self):
        if not self.fichier_path or not self.feuille_nom.get():
            return
        try:
            self.df = pd.read_excel(self.fichier_path, sheet_name=self.feuille_nom.get(), header=None)
            self.table.delete(*self.table.get_children())
            self.table["columns"] = list(range(len(self.df.columns)))
            for col in self.table["columns"]:
                self.table.heading(col, text=f"Col {col}")
                # Fixez la largeur à 50 pixels (ou autre valeur minimale souhaitée)
                self.table.column(col, width=100, minwidth=100)
            for i, row in self.df.head(20).iterrows():
                self.table.insert("", "end", values=list(row))
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur d'affichage : {e}")

    def ajouter_feuille(self):
        try:
            if not self.fichier_path or not self.feuille_nom.get() or not self.taille_entete_entry.get().isdigit():
                messagebox.showwarning("Champs manquants", "Veuillez renseigner le fichier, la feuille et la taille d'en-tête.")
                return
            entete = int(self.taille_entete_entry.get())
            feuille = Feuille(Fichier(self.fichier_path), self.feuille_nom.get(), entete)
            self.comparateur.ajouter_feuille(feuille)
            messagebox.showinfo("Ajout réussi", f"La feuille a été ajoutée au comparateur.")
            self.liste_feuilles.insert(tk.END, f"{feuille.fichier.nom} ({feuille.nom})")
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d'ajouter la feuille : {e}")
        

    def ajouter_feuille_dans_liste(self, feuille):
        self.comparateur.feuilles.append(feuille)
        self.liste_feuilles.insert(tk.END, f"{feuille.fichier.nom} ({feuille.nom})")

    def create_excel_preview_frame(self):
        self.excel_preview_frame = tk.LabelFrame(self, text="2. Aperçu du fichier Excel", bg="#f4f4f4")
        self.excel_preview_frame.pack(padx=10, pady=5, fill='both', expand=True)
        self.excel_preview_frame.pack_propagate(False)

         # Créer le Treeview directement dans le LabelFrame
        self.table = ttk.Treeview(self.excel_preview_frame, show='headings', height=5)
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

    def create_test_selector(self):
        
        self.test_frame = tk.LabelFrame(self, text="3. Sélection et exécution de tests statistiques", bg="#f4f4f4")
        self.test_frame.pack(fill="x", padx=10, pady=5)

        # Choix de la thématique
        tk.Label(self.test_frame, text="Type de test :").pack(side="left", padx=(5, 0))
        self.theme_var = tk.StringVar(value="Normalité")
        themes = ["Normalité", "Homogénéité des variances", "Comparaison de groupes", "Moyennes hebdomadaires"]
        self.theme_combo = ttk.Combobox(self.test_frame, values=themes, textvariable=self.theme_var, state="readonly", width=25)
        self.theme_combo.pack(side="left", padx=5)
        self.theme_combo.bind("<<ComboboxSelected>>", self.update_test_options)

        # Choix du test spécifique
        tk.Label(self.test_frame, text="Méthode :").pack(side="left", padx=(10, 0))
        self.test_method_var = tk.StringVar(value="shapiro")
        self.test_combo = ttk.Combobox(self.test_frame, values=["shapiro"], textvariable=self.test_method_var, state="readonly", width=20)
        self.test_combo.pack(side="left", padx=5)

        # Frame conditionnelle contenant une grille
        self.conditional_frame = tk.Frame(self.test_frame, bg="#f4f4f4")
        self.conditional_frame.pack(side="left", padx=10)

        # Sous-frame en grid pour les champs variables/groupe
        self.grid_frame = tk.Frame(self.conditional_frame, bg="#f4f4f4")
        self.grid_frame.pack()

        self.col_var_label = tk.Label(self.grid_frame, text="Variable :", bg="#f4f4f4")
        self.col_var = tk.Entry(self.grid_frame, width=15)
        self.col_var.insert(0, "Température")

        self.col_groupe_label = tk.Label(self.grid_frame, text="Groupe :", bg="#f4f4f4")
        self.col_groupe = tk.Entry(self.grid_frame, width=15)
        self.col_groupe.insert(0, "Salle")

        self.col_groupe1_label = tk.Label(self.grid_frame, text="Groupe 1 :", bg="#f4f4f4")
        self.col_groupe1 = tk.Entry(self.grid_frame, width=10)
        self.col_groupe1.insert(0, "A")

        self.col_groupe2_label = tk.Label(self.grid_frame, text="Groupe 2 :", bg="#f4f4f4")
        self.col_groupe2 = tk.Entry(self.grid_frame, width=10)
        self.col_groupe2.insert(0, "B")

        # Disposition en 2 lignes
        self.col_var_label.grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.col_var.grid(row=0, column=1, padx=5, pady=2)
        self.col_groupe_label.grid(row=0, column=2, sticky="w", padx=5, pady=2)
        self.col_groupe.grid(row=0, column=3, padx=5, pady=2)

        self.col_groupe1_label.grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.col_groupe1.grid(row=1, column=1, padx=5, pady=2)
        self.col_groupe2_label.grid(row=1, column=2, sticky="w", padx=5, pady=2)
        self.col_groupe2.grid(row=1, column=3, padx=5, pady=2)

        # Bouton d'exécution
        tk.Button(self.test_frame, text="Exécuter le test", command=self.executer_test_general).pack(side="left", padx=10)

        self.update_test_options()




    def update_test_options(self, event=None):
        theme = self.theme_var.get()

        if theme == "Normalité":
            options = ["shapiro", "dagostino", "anderson"]
            self.hide_conditional_fields()

        elif theme == "Homogénéité des variances":
            options = ["levene", "bartlett"]
            self.show_conditional_fields(show_groupes=False)

        elif theme == "Comparaison de groupes":
            options = ["student", "mannwhitney"]
            self.show_conditional_fields(show_groupes=True)

        elif theme == "Moyennes hebdomadaires":
            options = ["student", "mannwhitney"]
            self.show_conditional_fields(show_groupes=True)

        else:
            options = []
            self.hide_conditional_fields()

        self.test_combo["values"] = options
        self.test_combo.set(options[0] if options else "")




    def create_result_box(self):
        self.result_frame = tk.LabelFrame(self, text="4. Résultats du test", bg="#f4f4f4")
        self.result_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.result_text = tk.Text(self.result_frame, height=10, wrap="none")
        self.result_text.pack(fill="both", expand=True)

        scroll_y = tk.Scrollbar(self.result_frame, command=self.result_text.yview)
        scroll_y.pack(side="right", fill="y")
        self.result_text.configure(yscrollcommand=scroll_y.set)

    def executer_test(self):
        self.result_text.delete("1.0", tk.END)
        methode = self.methode_var.get()
        resultats = self.comparateur.tester_normalite(methode=methode)

        for col, res in resultats.items():
            if res["stat"] is None:
                self.result_text.insert(tk.END, f"Colonne {col} : données insuffisantes\n")
            else:
                normalite = "✅ Normale" if res["normal"] else "❌ Non normale"
                stat = f"{res['stat']:.4f}"
                pval = f"{res['p_value']:.4f}" if res["p_value"] else "—"
                self.result_text.insert(tk.END, f"{col} : stat={stat}, p={pval} → {normalite}\n")

    def show_conditional_fields(self, show_groupes=False):
        self.col_var_label.grid()
        self.col_var.grid()
        self.col_groupe_label.grid()
        self.col_groupe.grid()

        if show_groupes:
            self.col_groupe1_label.grid()
            self.col_groupe1.grid()
            self.col_groupe2_label.grid()
            self.col_groupe2.grid()
        else:
            self.col_groupe1_label.grid_remove()
            self.col_groupe1.grid_remove()
            self.col_groupe2_label.grid_remove()
            self.col_groupe2.grid_remove()




    def hide_conditional_fields(self):
        for widget in [
            self.col_var_label, self.col_var,
            self.col_groupe_label, self.col_groupe,
            self.col_groupe1_label, self.col_groupe1,
            self.col_groupe2_label, self.col_groupe2
        ]:
            widget.grid_remove()





    def executer_test_general(self):
        self.result_text.delete("1.0", tk.END)
        theme = self.theme_var.get()
        methode = self.test_method_var.get()

        if theme == "Normalité":
            resultats = self.comparateur.tester_normalite(methode=methode)
            for col, res in resultats.items():
                if res["stat"] is None:
                    self.result_text.insert(tk.END, f"Colonne {col} : données insuffisantes\n")
                else:
                    normalite = "✅ Normale" if res["normal"] else "❌ Non normale"
                    stat = f"{res['stat']:.4f}"
                    pval = f"{res['p_value']:.4f}" if res["p_value"] is not None else "—"
                    self.result_text.insert(tk.END, f"{col} : stat={stat}, p={pval} → {normalite}\n")

        elif theme == "Homogénéité des variances":
            var = self.col_var.get()
            groupe = self.col_groupe.get()
            groupe = to_int(groupe)
            var = to_int(var)

            resultats = self.comparateur.tester_homogeneite_variances({var: groupe}, methode=methode)
            if var in resultats:
                res = resultats[var]
                if res["stat"] is None:
                    self.result_text.insert(tk.END, f"Colonne {var} : données insuffisantes\n")
                else:
                    homog = "✅ Homogènes" if res["homogene"] else "❌ Variances différentes"
                    self.result_text.insert(tk.END, f"{var} : stat={res['stat']:.4f}, p={res['p_value']:.4f} → {homog}\n")

        elif theme == "Comparaison de groupes":
            var = self.col_var.get()
            groupe = self.col_groupe.get()
            # Si la colonne dans le DataFrame est de type chaîne, ne pas convertir en int
            groupe = to_int(groupe)
            var = to_int(var)

            groupe_1 = self.col_groupe1.get()
            groupe_2 = self.col_groupe2.get()

            # Si nécessaire, convertir en chaîne
            groupe_1 = str(groupe_1)
            groupe_2 = str(groupe_2)

            res = self.comparateur.tester_comparaison_groupes(var, groupe, groupe_1, groupe_2, methode=methode)

            if "error" in res:
                self.result_text.insert(tk.END, f"Erreur : {res['error']}\n")
            else:
                self.result_text.insert(tk.END,
                    f"{var} entre {res['groupe_1']} et {res['groupe_2']} ({methode}) :\n"
                    f"Stat={res['stat']:.4f}, p={res['p_value']:.4f} → {'✅ Différence significative' if res['significatif'] else '❌ Pas de différence'}\n")

        elif theme == "Moyennes hebdomadaires":
            var = self.col_var.get()
            groupe = self.col_groupe.get()
            # Si la colonne dans le DataFrame est de type chaîne, ne pas convertir en int
            # groupe = to_int(groupe)
            # var = to_int(var)

            groupe_1 = self.col_groupe1.get()
            groupe_2 = self.col_groupe2.get()

            # Si nécessaire, convertir en chaîne
            groupe_1 = str(groupe_1)
            groupe_2 = str(groupe_2)

            res = self.comparateur.tester_comparaison_groupes(var, groupe, groupe_1, groupe_2, methode=methode)

            if "error" in res:
                self.result_text.insert(tk.END, f"Erreur : {res['error']}\n")
            else:
                self.result_text.insert(tk.END,
                    f"{var} (moy. hebdo) entre {res['groupe_1']} et {res['groupe_2']} ({methode}) :\n"
                    f"Stat={res['stat']:.4f}, p={res['p_value']:.4f} → {'✅ Différence significative' if res['significatif'] else '❌ Pas de différence'}\n")


