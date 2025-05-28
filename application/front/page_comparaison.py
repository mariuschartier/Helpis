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
from structure.Selection_col import Selection_col



class ComparePage(tk.Frame):
    """ Page de comparaison de fichiers Excel pour effectuer des tests statistiques.
    Cette page permet de charger un fichier Excel, d'afficher un aperçu de son contenu,
    de sélectionner des tests statistiques et d'afficher les résultats.
    """
    def __init__(self, parent, controller):
        """ Initialise la page de comparaison.
        Args:
            parent (tk.Frame): Le parent de cette page.
            controller (Controller): Le contrôleur de l'application.
        """
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
        self.dico_groupe = {} 

        self.fonctions_courbes = [
        ("Normalité", self.tracer_courbe_normal),
        ("Q-Q plot", self.tracer_courbe_QQpolt),
        ]
        self.colonne_actuelle  = ""

        self.create_file_frame()

        self.create_excel_preview_frame()
        self.test_frame = tk.LabelFrame(self, text="3. Sélection et exécution de tests statistiques", bg="#f4f4f4")
        self.test_frame.pack(fill="x", padx=10, pady=5)
        self.create_result_box()
        self.create_result_tag()
        self.create_test_selector()


# frame de test ==========================================================
# Champ de chargement du fichier et de l'entete
    def create_file_frame(self):
        """ Crée le cadre pour le chargement du fichier Excel et la sélection de l'en-tête.
        """
        self.file_frame = tk.LabelFrame(self, text="1. Charger un fichier Excel", bg="#f4f4f4")
        self.file_frame.pack(fill="x", padx=10, pady=5)

        self.fichier_entry = tk.Entry(self.file_frame, width=80)
        self.fichier_entry.pack(side="left", padx=5, pady=5)

        tk.Button(self.file_frame, text="Parcourir", command=self.choisir_fichier).pack(side="left", padx=5)

        self.feuille_combo = ttk.Combobox(self.file_frame, textvariable=self.feuille_nom, state="readonly")
        self.feuille_combo.pack(side="left", padx=5)
        self.feuille_combo.bind("<<ComboboxSelected>>", lambda e: self.afficher_excel())


        # Choix de la taille de l'en-tête
        self.taille_entete_var = tk.StringVar()
        tk.Label(self.file_frame, text="Taille de l'en-tête :").pack(side="left", padx=(10, 0))
        self.taille_entete_entry = tk.Entry(self.file_frame, width=5,textvariable=self.taille_entete_var )
        self.taille_entete_var.set(1)  # Met à jour l'Entry avec 1
        self.taille_entete_var.trace_add("write", self.on_taille_entete_change)

        self.taille_entete_entry.pack(side="left", padx=5)
        tk.Button(self.file_frame, text="❓ Aide", command=self.ouvrir_aide).pack(side="right", padx=5)
        self.taille_entete_entry.bind("<KeyRelease>", self.on_key_release_int)
 
        tk.Button(self.file_frame, text="detail", command=self.ouvrir_popup_manipulation).pack(side="right", padx=5)

        # tk.Button(self.file_frame, text="Ajouter au comparateur", command=self.ajouter_feuille).pack(side="left", padx=10)
      
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


    def choisir_fichier(self):
        """ Ouvre une boîte de dialogue pour sélectionner un fichier Excel et charge la feuille sélectionnée.
        """
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return
        self.fichier_path = path
        self.fichier_entry.delete(0, tk.END)
        self.fichier_entry.insert(0, path)

        try:
            xls = pd.ExcelFile(path)
            self.feuille_combo["values"] = xls.sheet_names
            self.feuille_combo.set(xls.sheet_names[0])
            self.ajouter_feuille()
            self.afficher_excel()
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur de lecture du fichier : {e}")
               
    def maj_feuille(self):
        """ Met à jour la feuille avec les détails de la structure et l'ajoute au comparateur."""
        fichier = Fichier(self.fichier_path)
        
        self.comparateur.feuille.maj_feuille(fichier=fichier,
                                             nom=self.feuille_nom.get(),
                                             debut_data=self.details_structure["data_debut"],
                                             fin_data=self.details_structure["data_fin"]) 

        self.comparateur.feuille.entete.maj_entete(
                        entete_debut=self.details_structure["entete_debut"],
                        entete_fin=self.details_structure["entete_fin"],
                        nb_colonnes_secondaires=self.details_structure["nb_colonnes_secondaires"],
                        ligne_unite=self.details_structure["ligne_unite"],
                        structure=self.dico_structure)

        
    def ajouter_feuille(self):
        """ Ajoute la feuille sélectionnée au comparateur."""
        try:
            if not self.fichier_path or not self.feuille_nom.get() or not self.taille_entete_entry.get().isdigit():
                messagebox.showwarning("Champs manquants", "Veuillez renseigner le fichier, la feuille et la taille d'en-tête.")
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
            self.comparateur.ajouter_feuille(feuille)
            messagebox.showinfo("Ajout réussi", f"La feuille a été ajoutée au comparateur.")
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d'ajouter la feuille : {e}")
    
    def ouvrir_aide(self):
        """ Ouvre une fenêtre d'aide avec des instructions sur l'utilisation de l'application."""
        aide_popup = tk.Toplevel(self)
        aide_popup.title("Aide - Utilisation")
        aide_popup.geometry("600x400")

        texte = tk.Text(aide_popup, wrap="word", font=("Segoe UI", 10))
        texte.pack(fill="both", expand=True, padx=10, pady=10)

        contenu = (
            "🔍 Bienvenue dans l'application Testeur Excel\n\n"
            "Voici comment utiliser l'application :\n"

            "1 - Charger un fichier Excel :\n"
            "   - Cliquez sur le bouton 'Parcourir' pour sélectionner un fichier Excel.\n"
            "   - Sélectionnez la feuille à analyser dans le menu déroulant.\n"
            "   - Ajustez la taille de l'en-tête si nécessaire (par défaut 1).\n"
            "2 - Aperçu du fichier Excel :\n"
            "   - Un aperçu du fichier Excel s'affiche dans la zone prévue à cet effet.\n"
            "   - Vous pouvez faire défiler le tableau pour voir les données.\n"
            "3 - Sélection et exécution de tests statistiques :\n"
            "   - Choisissez le type de test statistique à exécuter dans le menu déroulant.\n"
            "   - Sélectionnez la méthode appropriée pour le test choisi.\n"
            "   - Si nécessaire, sélectionnez la variable et le groupe à analyser.\n"
            "   - Cliquez sur le bouton 'Exécuter le test' pour lancer l'analyse.\n"
            "4 - Résultats du test :\n"
            "   - Les résultats du test s'affichent dans la zone de résultats.\n"
            "   - Les résultats incluent la statistique du test, la valeur p et une indication de la significativité.\n"
            "   - Vous pouvez également tracer des courbes pour visualiser les données.\n\n"
        
        )
        
        texte.insert(tk.END, contenu)

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
            
        tk.Button(frame_cb, text="Réinitialisation", command=reset_valeur).pack(side="left", padx=10)


            

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

        tk.Button(frame_btns, text="✅ Appliquer", command=appliquer).pack(side="left", padx=10)
        tk.Button(frame_btns, text="❌ Annuler", command=popup.destroy).pack(side="left", padx=10)



# Affichage du fichier Excel dans le tableau ====================================================================================================================
    def afficher_excel(self):
        """Affiche le contenu du fichier Excel dans le tableau."""
        try:
            # Vider les anciennes données
            self.table.delete(*self.table.get_children())

            # Lire le fichier Excel
            self.df = pd.read_excel(self.fichier_path, sheet_name=self.feuille_nom.get(), header=None)
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
            self.colorier_ligne(self.details_structure["entete_debut"])
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de lire le fichier : {e}")
        self.dico_entete()

    def create_excel_preview_frame(self):
        """Crée le cadre pour l'aperçu du fichier Excel."""
        # Créer un LabelFrame
        self.excel_preview_frame = tk.LabelFrame(self, text="3. Aperçu du fichier Excel", bg="#f4f4f4")
        self.excel_preview_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Créer le Treeview avec une colonne pour les numéros de ligne
        self.table = ttk.Treeview(self.excel_preview_frame, show="tree headings", height=15)
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
        self.table.column("#0", width=40, minwidth=30, anchor="center", stretch=False)
        for name in col_names:
            self.table.heading(name, text=name)
            self.table.column(name, anchor="center", width=120, minwidth=100, stretch=True)

        # Exemple de remplissage avec numéros de ligne et valeurs fictives
        for i in range(50):
            values = [f"Valeur {j+1}" for j in range(nb_cols)]
            # Insérer avec le numéro de ligne (text=) et les valeurs
            self.table.insert("", "end", text=str(i + 1), values=values, tags=("ligne",))


        return self.excel_preview_frame

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






# SELECTION TESTS ====================================================================================================================
    def create_test_selector(self):
        """ Crée la section de sélection des tests statistiques."""
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
        self.var_selection = Selection_col(self.dico_structure)
        self.col_var = self.var_selection.get_frame_selection_grid( self.grid_frame,0,1)

        self.col_groupe_label = tk.Label(self.grid_frame, text="Groupe :", bg="#f4f4f4")
        self.groupe_selection = Selection_col(self.dico_structure)
        self.groupe_selection.action_selection = self.maj_selection_colonne
        self.col_groupe = self.groupe_selection.get_frame_selection_grid( self.grid_frame,0,3)

        # le dictionnaire correspond aux valeurs différentes de la colonne col_groupe
        self.col_groupe1_label = tk.Label(self.grid_frame, text="Groupe 1 :", bg="#f4f4f4")
        self.groupe1_selection = Selection_col(self.dico_groupe)
        self.col_groupe1 = self.groupe1_selection.get_frame_selection_grid( self.grid_frame,0,5)

        self.col_groupe2_label = tk.Label(self.grid_frame, text="Groupe 2 :", bg="#f4f4f4")
        self.groupe2_selection = Selection_col(self.dico_groupe)
        self.col_groupe2 = self.groupe2_selection.get_frame_selection_grid( self.grid_frame,1,5)

        # Disposition en 2 lignes
        self.col_var_label.grid(row=0, column=0, sticky="w", padx=5, pady=2)
        # self.col_var.grid(row=0, column=1, padx=5, pady=2)
        self.col_var_label.grid()

        self.col_groupe_label.grid(row=0, column=2, sticky="w", padx=5, pady=2)
        # self.col_groupe.grid(row=0, column=3, padx=5, pady=2)

        self.col_groupe1_label.grid(row=0, column=4, sticky="w", padx=5, pady=2)
        # self.col_groupe1.grid(row=1, column=1, padx=5, pady=2)
        self.col_groupe2_label.grid(row=1, column=4, sticky="w", padx=5, pady=2)
        # self.col_groupe2.grid(row=1, column=3, padx=5, pady=2)

        # Bouton d'exécution
        tk.Button(self.test_frame, text="Exécuter le test", command=self.executer_test_general).pack(side="left", padx=10)
        tk.Button(self.test_frame, text="afficher courbe variable", command=self.afficher_courbe_popup).pack(side="left", padx=10)


        self.update_test_options()

    def update_test_options(self, event=None):
        """ Met à jour les options de test en fonction du thème sélectionné."""
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

    def on_test_method_change(self,*args):
        """ Gère le changement de méthode de test sélectionnée."""
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

    def show_conditional_fields(self, show_groupes=False):
        """ Affiche les champs conditionnels en fonction du thème sélectionné."""
        self.col_var_label.grid()
        self.var_selection.grid()
        self.col_groupe_label.grid()
        self.groupe_selection.grid()

        if show_groupes:
            self.col_groupe1_label.grid()
            self.groupe1_selection.grid()
            self.col_groupe2_label.grid()
            self.groupe2_selection.grid()
        else:
            self.col_groupe1_label.grid_remove()
            self.groupe1_selection.grid_remove()
            self.col_groupe2_label.grid_remove()
            self.groupe2_selection.grid_remove()

    def hide_conditional_fields(self):
        """ Masque les champs conditionnels."""
        for widget in [
            
            self.col_groupe_label, self.groupe_selection,
            self.col_groupe1_label, self.groupe1_selection,
            self.col_groupe2_label, self.groupe2_selection
        ]:
            # self.col_var_label, self.var_selection,
            widget.grid_remove()

    def maj_selection_colonne(self):
        """ Met à jour les sélections de colonnes en fonction de la structure actuelle."""
        self.dico_colonne_groupe()
        self.var_selection.maj_donnees(self.dico_structure)
        self.groupe_selection.maj_donnees(self.dico_structure)
        self.groupe1_selection.maj_donnees(self.dico_groupe)
        self.groupe2_selection.maj_donnees(self.dico_groupe)

    def dico_colonne_groupe(self):
        """ Construit un dictionnaire des groupes à partir de la colonne sélectionnée."""
        self.dico_groupe = {}

        # Récupérer l’indice de la colonne correspondant au chemin sélectionné
        chemin_colonne = self.groupe_selection.chemin
        if chemin_colonne == "":
            return
        indice_colonne = self.comparateur.feuille.entete.placement_colonne.get(chemin_colonne)

        if indice_colonne is None:
            messagebox.showerror("Erreur", f"Colonne '{chemin_colonne}' non trouvée dans la feuille.")
            return

        for idx in range(self.comparateur.feuille.debut_data, self.comparateur.feuille.nb_ligne):
            data = self.df.iloc[idx, indice_colonne]
            if pd.isna(data):
                continue
            data = str(data)
            if data not in self.dico_groupe:
                self.dico_groupe[data] = {}  


    # Tracer des courbes 
    def tracer_courbe_normal(self, feuille,chemin=None):
        """ Trace une courbe normale pour la colonne sélectionnée."""
        try:
            feuille.entete.structure = self.dico_structure
            feuille.entete.placement_colonne = feuille.entete.set_position()
            incice_colonne = feuille.entete.placement_colonne[chemin]
            plot_histogram_normal(incice_colonne, feuille)
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors du traçage de la courbe : {e}")

    def tracer_courbe_QQpolt(self, feuille,chemin=None):
        """ Trace un Q-Q plot pour la colonne sélectionnée."""
        try:
            feuille.entete.structure = self.dico_structure
            feuille.entete.placement_colonne = feuille.entete.set_position()
            incice_colonne = feuille.entete.placement_colonne[chemin]
            plot_qqplot(incice_colonne, feuille)
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors du traçage de la courbe : {e}")

    def afficher_courbe_popup(self, event=None):
        """ Ouvre une fenêtre popup pour sélectionner et tracer une courbe."""
        # Récupère la feuille sélectionnée dans la liste
        selection = self.var_selection.chemin
        if not selection:
            messagebox.showwarning("Aucune sélection", "Veuillez sélectionner une feuille.")
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
        
        


        # Bouton pour tracer la courbe
        def valider():
            selection_courbe = listbox.curselection()
            if not selection_courbe:
                messagebox.showwarning("Aucune sélection", "Veuillez sélectionner une courbe.")
                return
            index_courbe = selection_courbe[0]
            _, fonction = self.fonctions_courbes[index_courbe]
            
            try:
                fonction(self.comparateur.feuille, selection)
                popup.destroy()
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors du tracé : {e}")

        btn_ok = tk.Button(popup, text="Tracer la courbe", command=valider)
        btn_ok.pack(pady=10)


        self.wait_window(popup)



# FRAME DE RESULTAT ====================================================================================================================
    def create_result_box(self):
        """ Crée la zone de résultats pour afficher les résultats des tests statistiques."""
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

    def append_text(self, new_content, color="black"):
        """ Ajoute du texte à la zone de résultats avec une couleur spécifique."""
        if not hasattr(self, "result_text"):
            print("Erreur : 'result_text' n'a pas été initialisé.")
            return
        # Créer le tag uniquement s'il n'existe pas
        if color not in self.result_text.tag_names():
            self.result_text.tag_config(color, foreground=color)
        # Insérer le texte avec le tag de couleur
        self.result_text.insert("end", new_content + "\n", color)
        self.result_text.see("end")

    def create_result_tag(self):
        """ Définit les tags de couleur pour la zone de résultats."""
        # Définir les tags une seule fois
        self.result_text.tag_config("black", foreground="black")
        self.result_text.tag_config("blue", foreground="blue")
        self.result_text.tag_config("red", foreground="red")
        self.result_text.tag_config("green", foreground="green")



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

            
            self.maj_feuille()
            self.maj_selection_colonne()
            return self.dico_structure

        except Exception as e:
            messagebox.showerror("Erreur",  f"Fichier et taille d'entete requis.{e}")
            # messagebox.showerror("Erreur", f"Impossible de construire le dictionnaire d'en-tête : {e}")
            return {}




#EXECUTION DES TESTS ====================================================================================================================

    def executer_test_general(self):
        """ Exécute le test statistique sélectionné et affiche les résultats."""
        

        theme = self.theme_var.get()
        methode = self.test_method_var.get()
        
        # print("colonne_actuelle "+self.var_selection.chemin)
        if not self.var_selection.chemin:
            messagebox.showwarning("Aucune sélection", "Veuillez sélectionner une colonne.")
            return
        
        if theme == "Normalité":
            resultats = self.comparateur.tester_normalite(self.var_selection.chemin, methode=methode)
            if resultats["stat"] is None:
                self.append_text( f"Colonne {self.self.var_selection.chemin} : données insuffisantes\n")
            else:
                normalite = "✅ Normale" if resultats["normal"] else "❌ Non normale"
                stat = f"{resultats['stat']:.4f}"
                pval = f"{resultats['p_value']:.4f}" if resultats["p_value"] is not None else "—"
                self.append_text( f"{self.var_selection.chemin} : stat={stat}, p={pval} → {normalite}\n")

        elif theme == "Homogénéité des variances":
            var = self.var_selection.chemin
            groupe = self.groupe_selection.chemin

            i_groupe = self.comparateur.feuille.entete.placement_colonne[groupe]
            i_var = self.comparateur.feuille.entete.placement_colonne[var]


            # print(f"i_groupe : {groupe}, i_var : {var}")
            res = self.comparateur.tester_homogeneite_variances(var, groupe, methode=methode)
            # print(res)
            if res["stat"] is None:
                self.append_text(f"Colonne {var} : données insuffisantes\n")
            else:
                homog = "✅ Homogènes" if res["homogene"] else "❌ Variances différentes"
                self.append_text(f"{var} : stat={res['stat']:.4f}, p={res['p_value']:.4f} → {homog}\n")


        elif theme == "Comparaison de groupes":

            var = self.var_selection.chemin
            groupe = self.groupe_selection.chemin

            groupe_1 = self.groupe1_selection.chemin
            groupe_2 = self.groupe2_selection.chemin

            # 🔸 Appel à la méthode avec les bons types
            res = self.comparateur.tester_comparaison_groupes(
                variable=var,
                groupe=groupe,
                groupe_1=str(groupe_1),
                groupe_2=str(groupe_2),
                methode=methode
            )

            if res.get("error"):
                self.append_text(f"❌ Erreur : {res['error']}\n")
            else:
                self.append_text(
                    f"{var} entre {res['groupe_1']} et {res['groupe_2']} ({methode}) :\n"
                    f"Stat = {res['stat']:.4f}, p = {res['p_value']:.4f} → "
                    f"{'✅ Différence significative' if res['significatif'] else '❌ Pas de différence'}\n"
                )


        elif theme == "Moyennes hebdomadaires":
            var = self.var_selection.chemin
            groupe = self.groupe_selection.chemin

            groupe_1 = self.groupe1_selection.chemin
            groupe_2 = self.groupe2_selection.chemin

            # 🔸 Appel à la méthode avec les bons types
            res = self.comparateur.tester_comparaison_moyennes_hebdo(
                variable=var,
                groupe=groupe,
                groupe_1=str(groupe_1),
                groupe_2=str(groupe_2),
                methode=methode
            )

            if "error" in res:
                self.append_text(f"Erreur : {res['error']}\n")
            else:
                self.append_text(                    f"{var} (moy. hebdo) entre {res['groupe_1']} et {res['groupe_2']} ({methode}) :\n"
                    f"Stat={res['stat']:.4f}, p={res['p_value']:.4f} → {'✅ Différence significative' if res['significatif'] else '❌ Pas de différence'}\n")


