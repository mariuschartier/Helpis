from tkinter import filedialog, messagebox, ttk,simpledialog
import tkinter as tk
import pandas as pd

from tests.ComparateurFichiers import ComparateurFichiers
from tests.fonctions import to_int
from tests.courbes import plot_histogram_normal,plot_qqplot
import numpy as np
import matplotlib.pyplot as plt
import scipy.stats as stats

from structure.Entete import Entete
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
        self.details_structure = {
            "entete_debut": 0,
            "entete_fin": 0,
            "data_debut": 1,
            "data_fin": None,
            "nb_colonnes_secondaires": 0,
            "ligne_unite": 0,
            "ignorer_vide": True
        }
        self.dico_structure =  {}
        self.fonctions_courbes = [
        ("Normalité", self.tracer_courbe_normal),
        ("Q-Q plot", self.tracer_courbe_QQpolt),
        ]
        

        self.create_file_frame()
        self.block_liste_feuille()

        self.create_excel_preview_frame()
        
        
        self.test_frame = tk.LabelFrame(self, text="3. Sélection et exécution de tests statistiques", bg="#f4f4f4")
        self.test_frame.pack(fill="x", padx=10, pady=5)
        self.create_result_box()
        self.create_result_tag()
        self.create_test_selector()


# frame de test ==========================================================
    def create_file_frame(self):
        self.file_frame = tk.LabelFrame(self, text="1. Charger un fichier Excel", bg="#f4f4f4")
        self.file_frame.pack(fill="x", padx=10, pady=5)

        self.fichier_entry = tk.Entry(self.file_frame, width=80)
        self.fichier_entry.pack(side="left", padx=5, pady=5)

        tk.Button(self.file_frame, text="Parcourir", command=self.choisir_fichier).pack(side="left", padx=5)

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

        tk.Button(self.file_frame, text="Ajouter au comparateur", command=self.ajouter_feuille).pack(side="left", padx=10)

        
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
        self.liste_frame = tk.LabelFrame(self, text="Fichiers ajoutés", bg="#f4f4f4")
        self.liste_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Cadre pour organiser la Listbox et le bouton
        self.list_bouton_frame = tk.Frame(self.liste_frame)
        self.list_bouton_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # La Listbox
        self.liste_feuilles = tk.Listbox(self.list_bouton_frame, height=5, selectmode="extended")
        self.liste_feuilles.pack(side="left", fill="both", expand=True)

        #scrollbar
        scrollbar_list = tk.Scrollbar(self.list_bouton_frame, command=self.liste_feuilles.yview)
        scrollbar_list.pack(side="right", fill="y")
        self.liste_feuilles.config(yscrollcommand=scrollbar_list.set)
        # Ajouter un bouton pour afficher la courbe
        self.liste_feuilles.bind("<Double-Button-1>", self.afficher_courbe_popup)

        # Le bouton de suppression à droite
        tk.Button(self.list_bouton_frame, text="Retirer le test sélectionné", command=self.supprimer_test).pack(side="right", pady=5)


    def supprimer_test(self):
        selection = self.liste_feuilles.curselection()
        if not selection:
            return
    
        # Supprimer dans l'ordre inverse pour éviter les décalages d'index
        for index in reversed(selection):
            self.liste_feuilles.delete(index)
    
        self.append_text(f"{len(selection)} test(s) supprimé(s).\n")

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
        
        # self.test_frame = tk.LabelFrame(self, text="3. Sélection et exécution de tests statistiques", bg="#f4f4f4")
        # self.test_frame.pack(fill="x", padx=10, pady=5)

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
        self.test_method_var.trace("w", self.on_test_method_change)

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
        # Cadre pour les résultats du test
        self.result_frame = tk.LabelFrame(self, text="4. Résultats du test", bg="#f4f4f4")
        self.result_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Barre de défilement verticale
        scroll_y = tk.Scrollbar(self.result_frame, orient="vertical")
        scroll_y.pack(side="right", fill="y")

        # Barre de défilement horizontale
        scroll_x = tk.Scrollbar(self.result_frame, orient="horizontal")
        scroll_x.pack(side="bottom", fill="x")

        # Zone de texte avec padding
            # Zone de texte avec padding
        self.result_text = tk.Text(
            self.result_frame, 
            height=10, 
            wrap="none",  # Pas de retour à la ligne automatique
            xscrollcommand=scroll_x.set, 
            yscrollcommand=scroll_y.set,
            padx=5, 
            pady=5
        )
        self.result_text.pack(fill="both", expand=True)

        # Méthode pour ajouter du texte sans le remplacer
 
        self.result_text.pack(fill="both", expand=True)

        # Configuration des barres de défilement
        scroll_y.config(command=self.result_text.yview)
        scroll_x.config(command=self.result_text.xview)


# petite fonction =================================================================
    def create_result_tag(self):
        # Définir les tags une seule fois
        self.result_text.tag_config("black", foreground="black")
        self.result_text.tag_config("blue", foreground="blue")
        self.result_text.tag_config("red", foreground="red")
        self.result_text.tag_config("green", foreground="green")

    def append_text(self, new_content, color="black"):
        if not hasattr(self, "result_text"):
            print("Erreur : 'result_text' n'a pas été initialisé.")
            return
        # Créer le tag uniquement s'il n'existe pas
        if color not in self.result_text.tag_names():
            self.result_text.tag_config(color, foreground=color)
        # Insérer le texte avec le tag de couleur
        self.result_text.insert("end", new_content + "\n", color)
        self.result_text.see("end")

    def on_key_release(self, event):
        if not self.taille_entete_entry.get().isdigit() and self.taille_entete_entry.get() != "":
            messagebox.showwarning("Validation", "Veuillez entrer un nombre entier.")
            self.taille_entete_entry.delete(0, tk.END)

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


    # Exemple de fonctions pour tracer des courbes
    def tracer_courbe_normal(self, feuille,chemin=None):
        try:
            feuille.entete.structure = self.dico_structure
            feuille.entete.placement_colonne = feuille.entete.set_position()
            incice_colonne = feuille.entete.placement_colonne[chemin]
            plot_histogram_normal(incice_colonne, feuille)
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors du traçage de la courbe : {e}")

    def tracer_courbe_QQpolt(self, feuille,chemin=None):
        try:
            feuille.entete.structure = self.dico_structure
            feuille.entete.placement_colonne = feuille.entete.set_position()
            incice_colonne = feuille.entete.placement_colonne[chemin]
            plot_qqplot(incice_colonne, feuille)
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors du traçage de la courbe : {e}")


    def afficher_courbe_popup(self, event=None):
        # Récupère la feuille sélectionnée dans la liste
        selection = self.liste_feuilles.curselection()
        if not selection:
            messagebox.showwarning("Aucune sélection", "Veuillez sélectionner une feuille.")
            return

        index = selection[0]
        feuille_obj = self.comparateur.feuilles[index]
        if not feuille_obj:
            messagebox.showerror("Erreur", "Feuille non trouvée.")
            return

        # Crée une nouvelle fenêtre popup
        popup = tk.Toplevel(self)
        popup.title("Choisissez la courbe à afficher")
        popup.geometry("400x300")
        popup.grab_set()

        tk.Label(popup, text="Sélectionnez la courbe à tracer :").pack(pady=10)

        listbox = tk.Listbox(popup, height=6)
        for nom, _ in self.fonctions_courbes:
            listbox.insert(tk.END, nom)
        listbox.pack(fill="both", expand=True, padx=10, pady=5)

        # Appel pour créer la partie de sélection de colonnes
        
        get_path = self.select_column_path(popup)


        # Bouton pour tracer la courbe
        def valider():
            selection_courbe = listbox.curselection()
            if not selection_courbe:
                messagebox.showwarning("Aucune sélection", "Veuillez sélectionner une courbe.")
                return

            chemin_1 = get_path()  # On appelle get_path() maintenant, après la sélection.
            if not chemin_1:
                messagebox.showerror("Erreur", "Veuillez sélectionner une colonne cible.")
                return

            index_courbe = selection_courbe[0]
            _, fonction = self.fonctions_courbes[index_courbe]
            
            try:
                fonction(feuille_obj, chemin_1)
                popup.destroy()
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors du tracé : {e}")

        btn_ok = tk.Button(popup, text="Tracer la courbe", command=valider)
        btn_ok.pack(pady=10)


        self.wait_window(popup)



    def select_column_path(self, popup):
        """
        Crée une interface pour sélectionner une colonne et ses sous-catégories, et renvoie le chemin sélectionné.

        Args:
        - dico_structure (dict): Dictionnaire représentant la structure hiérarchique des colonnes.
        - popup (tk.Toplevel): Fenêtre popup où les combobox seront placées.

        Returns:
        - str: Chemin complet sélectionné.
        """
        frame = tk.Frame(popup)
        frame.pack()

        # Choix de la colonne principale
        colonne_combo = ttk.Combobox(frame, values=list(self.dico_structure.keys()), state="readonly")
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
            if selected_col in self.dico_structure:
                add_combobox(frame, 1, self.dico_structure[selected_col], comboboxes)

        colonne_combo.bind("<<ComboboxSelected>>", on_colonne_selection)

        def get_path():
            col1 = colonne_combo.get()
            selection = [combo.get() for combo, _ in comboboxes if combo.get()]

            return " > ".join([col1] + selection) if col1 else None
        
        return get_path

#EXECUTION DES TESTS ==========================================================

    def on_test_method_change(self,*args):
        selected_method = self.test_method_var.get()
        self.append_text(f"Méthode sélectionnée : {selected_method}", color="blue")
        dico_methode_contrainte = {
            "shapiro": "✅ Taille de l’échantillon : 3 ≤ n ≤ 2000\n"
            "              ✅ Données quantitatives continues.\n",

            "dagostino": "✅ Taille de l’échantillon : n ≥ 20.\n"
            "              ✅ Données quantitatives continues.\n",

            "anderson": "✅ Aucune limite stricte sur n, mais plus précis pour n ≥ 50.\n"
            "              ✅ Données quantitatives continues.\n",

            "levene": "✅ Pas de normalité requise.\n"
            "              ✅ Groupes indépendants.\n",

            "bartlett": "✅ Les données doivent être normales.\n"
            "              ✅ Groupes indépendants.\n",

            "student": "✅ Données normales dans chaque groupe.\n"
            "              ✅ Homogénéité des variances.\n"
            "              ✅ Groupes indépendants ou appariés.\n",
            "mannwhitney": "✅ Aucune condition de normalité requise.\n"
            "              ✅ Données ordinales ou continues.\n"
            "              ✅ Groupes indépendants.\n"
        }
        self.append_text(f"Contraintes : {dico_methode_contrainte[selected_method]}", color="red")

    def executer_test(self):
        methode = self.methode_var.get()
        resultats = self.comparateur.tester_normalite(methode=methode)

        for col, res in resultats.items():
            if res["stat"] is None:
                self.append_text( f"Colonne {col} : données insuffisantes\n")
            else:
                normalite = "✅ Normale" if res["normal"] else "❌ Non normale"
                stat = f"{res['stat']:.4f}"
                pval = f"{res['p_value']:.4f}" if res["p_value"] else "—"
                self.append_text( f"{col} : stat={stat}, p={pval} → {normalite}\n")

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
        theme = self.theme_var.get()
        methode = self.test_method_var.get()

        if theme == "Normalité":
            resultats = self.comparateur.tester_normalite(methode=methode)
            for col, res in resultats.items():
                if res["stat"] is None:
                    self.append_text( f"Colonne {col} : données insuffisantes\n")
                else:
                    normalite = "✅ Normale" if res["normal"] else "❌ Non normale"
                    stat = f"{res['stat']:.4f}"
                    pval = f"{res['p_value']:.4f}" if res["p_value"] is not None else "—"
                    self.append_text( f"{col} : stat={stat}, p={pval} → {normalite}\n")

        elif theme == "Homogénéité des variances":
            var = self.col_var.get()
            groupe = self.col_groupe.get()
            groupe = to_int(groupe)
            var = to_int(var)

            resultats = self.comparateur.tester_homogeneite_variances({var: groupe}, methode=methode)
            if var in resultats:
                res = resultats[var]
                if res["stat"] is None:
                    self.append_text(f"Colonne {var} : données insuffisantes\n")
                else:
                    homog = "✅ Homogènes" if res["homogene"] else "❌ Variances différentes"
                    self.append_text(f"{var} : stat={res['stat']:.4f}, p={res['p_value']:.4f} → {homog}\n")

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
                self.append_text(f"Erreur : {res['error']}\n")
            else:
                self.append_text(                    f"{var} entre {res['groupe_1']} et {res['groupe_2']} ({methode}) :\n"
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
                self.append_text(f"Erreur : {res['error']}\n")
            else:
                self.append_text(                    f"{var} (moy. hebdo) entre {res['groupe_1']} et {res['groupe_2']} ({methode}) :\n"
                    f"Stat={res['stat']:.4f}, p={res['p_value']:.4f} → {'✅ Différence significative' if res['significatif'] else '❌ Pas de différence'}\n")


