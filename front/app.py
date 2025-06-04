import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import jsonpickle


from structure.Fichier import Fichier
from back.recherche_erreur.Test_gen import Test_gen
from back.recherche_erreur.Test_spe import Test_spe
from structure.Feuille import Feuille
from structure.Entete import Entete
from structure.Selection_col import Selection_col

import os
import json
from pathlib import Path
import ttkbootstrap as ttkb
from ttkbootstrap.constants import *
import threading

import webbrowser



class ExcelTesterApp(ttkb.Frame):
    """Page pour detecter les erreurs avec des tests des fichiers Excel."""
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
            "entete_fin": 0,
            "data_debut": 1,
            "data_fin": None,
            "nb_colonnes_secondaires": 0,
            "ligne_unite": 0,
            "ignorer_vide": True
        }
        

        # === Canvas + Scroll principal ===
        self.canvas = tk.Canvas(self, bg="#f4f4f4")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.scroll_y = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scroll_y.pack(side="right", fill="y")

        # Créer le frame qui sera scrollable
        self.scrollable_frame = tk.Frame(self.canvas, bg="#f4f4f4")

        # Lier le resize du frame à la zone de scroll
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )

        # Créer une fenêtre dans le canvas pour y placer le frame
        self.canvas_window = self.canvas.create_window(
            (0, 0), window=self.scrollable_frame, anchor="nw"
        )

        # Assurer que la largeur du frame s'ajuste à celle du canvas
        def resize_frame(event):
            canvas_width = event.width
            self.canvas.itemconfig(self.canvas_window, width=canvas_width)

        self.canvas.bind("<Configure>", resize_frame)

        # Relier la scrollbar au canvas
        self.canvas.configure(yscrollcommand=self.scroll_y.set)

        # Mise à jour des tâches d'idéalisation
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
        
        self.desactivation_bouton()
        
        self.register_scrollable_widgets()
        
        self._bind_mousewheel_to_widget(self.test_listbox)
        self._bind_mousewheel_to_widget(self.result_text)
        self._bind_mousewheel_to_widget(self.table)
        self._bind_mousewheel_to_widget(self.erreur_table)
        


# Gestion du scroll =========================================================================================================
    
    def register_scrollable_widgets(self):
        """Enregistre les widgets scrollables et lie le scroll de la souris."""
        scrollables = [
            self.test_listbox,
            self.result_text,
            self.table,
            self.erreur_table,
        ]
    
        for widget in scrollables:
                self._bind_mousewheel_to_widget(widget)
                
    def _disable_scroll_on_combo(self, widget):
        """Désactive le scroll de la souris sur un widget spécifique."""
        widget.bind("<Enter>", lambda e: self.canvas.unbind_all("<MouseWheel>"))
        widget.bind("<Leave>", lambda e: self.canvas.bind_all("<MouseWheel>", self._on_mousewheel))
    
    def _bind_mousewheel_to_widget(self, widget):
        """Lie le scroll de la souris à un widget spécifique."""
        widget.bind("<Enter>", lambda e: self._set_active_scroll_target(widget))
        widget.bind("<Leave>", lambda e: self._set_active_scroll_target(None))
    
    def _set_active_scroll_target(self, widget):
        """Définit le widget actif pour le scroll de la souris."""
        self._active_mouse_scroll_target = widget
    
    def _on_mousewheel(self, event):
        """Gère le scroll de la souris pour les widgets enregistrés."""
        target = getattr(self, "_active_mouse_scroll_target", None)
    
        if isinstance(target, (tk.Text, tk.Listbox, ttk.Treeview)):
            target.yview_scroll(int(-1 * (event.delta / 120)), "units")
        else:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")




# Frame =========================================================================================================
    
    def create_file_frame(self):
        """Crée le cadre pour charger le fichier Excel et configurer l'en-tête avec wrapping dynamique et taille minimale."""
        self.file_frame = tk.LabelFrame(self.scrollable_frame, text="1. Charger un fichier Excel", bg="#f4f4f4")
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

        aide_btn = ttkb.Button(self.file_frame, text="❓ Aide", command=self.ouvrir_aide, width=10)
        self.widgets_file_frame.append(aide_btn)



        self.file_frame.bind("<Configure>", lambda event: self.arrange_widgets_file_frame(self.file_frame, self.widgets_file_frame))

        return self.file_frame

    def arrange_widgets_file_frame(self, container:tk.Frame, widgets, event=None):
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
        if self.details_structure["data_debut"] <= self.details_structure["entete_fin"]:
            self.details_structure["data_debut"] = self.details_structure["entete_fin"]+1

        self.enlever_toutes_couleurs()
        self.colorier_lignes_range(
            self.details_structure["entete_debut"],
            self.details_structure["entete_fin"])

        self.dico_entete()
 
    def ouvrir_aide(self):
        """Ouvre une fenêtre d'aide avec des instructions sur l'utilisation de l'application."""
        aide_popup = tk.Toplevel(self)
        aide_popup.title("Aide - Utilisation")
        aide_popup.geometry("600x400")
        aide_popup.grab_set()

        texte = tk.Text(aide_popup, wrap="word", font=("Segoe UI", 10))
        texte.pack(fill="both", expand=True, padx=10, pady=10)

        contenu = (
                    "🔍 Bienvenue sur la page de vérification Excel\n\n"
                    "Voici comment utiliser l'application :\n"
                    "1 - Cliquez sur le bouton 'Parcourir' pour charger un fichier Excel (.xlsx).\n"
                    "2 - Choisissez la feuille à analyser dans la liste déroulante.\n"
                    "3 - Indiquez la taille de l’en-tête (nombre de lignes au début du tableau).\n"
                    "   - Vous pouvez utiliser le bouton 'Détail' pour accéder à une gestion plus précise de l'en-tête.\n"
                    "4 - Ajoutez un test générique (valeur minimale, maximale ou entre) ou un test spécifique.\n"
                    "5 - Cliquez sur le bouton 'Exécuter les tests' pour analyser le fichier.\n\n"
                    "💡 Les erreurs sont colorées dans le fichier Excel et listées dans la zone de résultats.\n"
            "📌 Vous pouvez faire défiler l’aperçu et les erreurs avec les barres de défilement.\n"
        )
        
        texte.insert(tk.END, contenu)
   
    def choisir_fichier(self):
        """Ouvre un dialogue pour choisir un fichier Excel et charge les feuilles disponibles."""
        filepath = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx")],
            initialdir="sauvegardes/results",  # Dossier de fichiers à convertir
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
                self.lien_fichier()
                self.activation_bouton()
            except Exception as e:
                messagebox.showerror("Erreur", f"Impossible de lire les feuilles du fichier : {e}")



    def lien_fichier(self):
        """Crée un lien cliquable pour ouvrir le fichier, en supprimant le précédent si nécessaire."""
        # Si le label existe déjà, le détruire avant d'en créer un nouveau
        if hasattr(self, 'link_label') and self.link_label.winfo_exists():
            self.link_label.destroy()

        # Créer un nouveau label
        self.link_label = tk.Label(self.results_frame, text="Cliquez ici pour ouvrir le fichier",
            fg="blue", cursor="hand2", font=("Arial", 10, "underline"))
        self.link_label.pack(padx=10, pady=10)

        # Bind le clic à la fonction d'ouverture
        self.link_label.bind("<Button-1>", lambda e: self.ouvrir_fichier(self.fichier_path))

    def ouvrir_fichier(self,chemin):
        """Ouvre le fichier Excel dans le navigateur ou l'application par défaut."""
        # Chemin vers le fichier
        fichier = chemin
        # Vérifier si le fichier existe, sinon ouvrir via le navigateur
        if os.path.exists(fichier):
            os.startfile(fichier)  # Sur Windows
        else:
            # Si pas Windows ou si vous souhaitez ouvrir dans le navigateur :
            webbrowser.open(fichier)
    
        # Ouvrir le popup de manipulation de l'entete detaillée

    def ouvrir_popup_manipulation(self):
        """Ouvre un popup pour configurer les paramètres avancés de la feuille."""

        print("details_structure :")
        print(self.details_structure)

        if self.df is None:
            messagebox.showerror("Erreur", "Un fichier doit être sélectionné.")
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
            ("Ligne des unités :", "ligne_unite"),
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

            if key == "data_fin" and self.details_structure["data_fin"] is None:
                try:
                    valeur_defaut = str(self.df.shape[0])
                except AttributeError:
                    valeur_defaut = ""
            else:
                valeur_defaut = valeurs_par_defaut.get(key, "")

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
            print("valeur :")
            print(valeurs)
            print("details_structure :")
            print(self.details_structure)

            if hasattr(self, "taille_entete_entry"):
                self.taille_entete_entry.delete(0, tk.END)
                self.taille_entete_entry.insert(0, str(taille_entete))

            popup.destroy()

        ttkb.Button(frame_btns, text="✅ Appliquer", command=appliquer_parametres).pack(side="left", padx=10)
        ttkb.Button(frame_btns, text="❌ Annuler", command=popup.destroy).pack(side="left", padx=10)


# Champs de test
    def create_test_buttons_frame(self):
        """Crée le cadre pour les boutons de test et organise leur disposition."""
        self.frame_btn_test = tk.Frame(self.scrollable_frame)
        self.frame_btn_test.pack(fill="both", expand=True, padx=10, pady=5)


        # Créer les boutons et les stocker dans une liste
        boutons_tests = []

        self.btn_popup_ajouter_test_gen = ttkb.Button(self.frame_btn_test, text="Ajouter un test générique", command=self.controller.bind_button(self.popup_ajouter_test_gen))
        # btn1.pack(side="left", padx=10)
        boutons_tests.append(self.btn_popup_ajouter_test_gen)

        self.btn_popup_ajouter_test_spe = ttkb.Button(self.frame_btn_test, text="Ajouter un test spécifique", command=self.controller.bind_button(self.popup_ajouter_test_spe))
        # btn2.pack(side="left", padx=10)
        boutons_tests.append(self.btn_popup_ajouter_test_spe)

        self.btn_executer_tests = ttkb.Button(self.frame_btn_test, text="Exécuter les tests", command=self.controller.bind_button(self.executer_tests))
        # btn3.pack(side="left", padx=10)
        boutons_tests.append(self.btn_executer_tests)

        self.btn_sauvegarder_tests = ttkb.Button(self.frame_btn_test, text="💾 Sauvegarder les tests", command=self.controller.bind_button(self.sauvegarder_tests))
        # btn4.pack(side="left", padx=10)
        boutons_tests.append(self.btn_sauvegarder_tests)

        self.btn_importer_tests = ttkb.Button(self.frame_btn_test, text="📂 Importer des tests", command=self.controller.bind_button(self.importer_tests))
        # btn5.pack(side="left", padx=10)
        boutons_tests.append(self.btn_importer_tests)

        # Appliquer la fonction pour organiser les boutons
        # self.arrange_widgets_file_frame(self.frame_btn_test, boutons_tests)
        self.frame_btn_test.bind("<Configure>", lambda event: self.arrange_widgets_file_frame(self.frame_btn_test, boutons_tests))


        return self.frame_btn_test
    
    def create_test_list_frame(self):
        """Crée le cadre pour la liste des tests."""
        self.test_list_frame = tk.LabelFrame(self.scrollable_frame, text="2. Liste des tests", bg="#f4f4f4")
        self.test_list_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Frame pour la Listbox et la scrollbar côte à côte
        list_frame = tk.Frame(self.test_list_frame, bg="#f4f4f4")
        list_frame.pack(fill="both", expand=True)

        # Listbox
        self.test_listbox = tk.Listbox(list_frame, height=5, selectmode="extended")
        self.test_listbox.pack(side="left", fill="both", expand=True)

        # Scrollbar attachée à la Listbox
        scrollbar_list = tk.Scrollbar(list_frame, command=self.test_listbox.yview)
        scrollbar_list.pack(side="right", fill="y")
        self.test_listbox.config(yscrollcommand=scrollbar_list.set)

        # Double clic sur la liste
        self.test_listbox.bind("<Double-Button-1>", self.afficher_details_popup)

        # Bouton pour retirer le test sélectionné
        ttkb.Button(self.test_list_frame, text="Retirer le test sélectionné", command=self.supprimer_test).pack(pady=5)

        return self.test_list_frame

    def supprimer_test(self):
        """Supprime le test sélectionné dans la liste des tests."""
        selection = self.test_listbox.curselection()
        if not selection:
            return
    
        # Supprimer dans l'ordre inverse pour éviter les décalages d'index
        for index in reversed(selection):
            del self.tests[index]
        self.append_text( f"{len(selection)} test(s) supprimé(s).", color="red")
        self.test_listbox.delete(*self.test_listbox.curselection())
    
    def sauvegarder_tests(self):
        """Sauvegarde les tests dans un fichier JSON."""
        Path("sauvegardes_tests").mkdir(exist_ok=True)
        
        chemin = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
            initialdir="sauvegardes/sauvegardes_tests",
            title="Sauvegarder les tests"
        )
        
        if not chemin:
            return

        export = []

        for test in self.tests:
            obj = test[0]

            if isinstance(obj, Test_gen):
                _, type_test, val_min, val_max = test
                export.append({
                    "type": "gen",
                    "nom": obj.nom,
                    "critere": obj.critere,
                    "test_type": type_test,
                    "val_min": val_min,
                    "val_max": val_max
                })

            elif isinstance(obj, Test_spe):
                # test, type_selected, chemin_1, chemin_2, val1, val2
                _, test_type, col1, col2, val1, val2 = test
                export.append({
                    "type": "spe",
                    "nom": obj.nom,
                    "test_type": test_type,
                    "col1": col1,
                    "col2": col2,
                    "val1": val1,
                    "val2": val2
                })

            else:
                # Par défaut, utiliser jsonpickle pour les objets inconnus
                export.append({
                    "type": "unknown",
                    "data": jsonpickle.encode(obj)
                })

        # Sauvegarder en JSON
        with open(chemin, "w", encoding="utf-8") as f:
            json.dump(export, f, ensure_ascii=False, indent=2)      
    
    def importer_tests(self):
        """Ouvre un dialogue pour importer des tests depuis un fichier JSON."""
        chemin = filedialog.askopenfilename(
            filetypes=[("JSON", "*.json")],
            initialdir="sauvegardes/sauvegardes_tests",
            title="Importer un fichier de tests"
        )

        if not chemin:
            return

        try:
            with open(chemin, "r", encoding="utf-8") as f:
                data = json.load(f)

            for test_data in data:
                try:
                    # Type de test
                    test_type = test_data.get("type")

                    if test_type == "gen":
                        # Créer une instance de Test_gen
                        obj = Test_gen(
                            nom=test_data["nom"],
                            critere=test_data["critere"]
                        )
                        # Ajouter à la liste des tests
                        self.tests.append((
                                obj, 
                                test_data["test_type"], 
                                test_data.get("val_min"), 
                                test_data.get("val_max")
                            ))
                        self.test_listbox.insert(tk.END, f"[GEN] {test_data['nom']} ({test_data['test_type']})")

                    elif test_type == "spe":
                        # Créer une instance de Test_spe
                        feuille = None  # Si nécessaire, adapter pour utiliser un objet ou un fichier
                        obj = Test_spe(
                            nom=test_data["nom"],
                            feuille=feuille
                        )
                        # Ajouter à la liste des tests
                        self.tests.append((
                            obj,
                            test_data["test_type"],
                            test_data["col1"],
                            test_data["col2"],
                            test_data.get("val1"),
                            test_data.get("val2")
                        ))
                        self.test_listbox.insert(tk.END, f"[SPE] {test_data['nom']} ({test_data['test_type']})")

                    elif test_type == "unknown":
                        # Décoder l'objet inconnu avec jsonpickle
                        obj = jsonpickle.decode(test_data["data"])
                        self.tests.append((obj,))
                        self.test_listbox.insert(tk.END, f"[UNKNOWN] {type(obj).__name__}")

                    else:
                        print(f"Type de test inconnu : {test_type}")

                except Exception as e:
                    print(f"Erreur lors du chargement d'un test : {e}")

        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de charger les tests : {e}")
        print(self.tests)

    def afficher_details_popup(self, event):
        """Affiche un popup avec les détails du test sélectionné dans la liste."""
        selection = self.test_listbox.curselection()
        if not selection:
            return

        index = selection[0]
        test_info = self.tests[index]

        popup = tk.Toplevel(self)
        popup.title("Détails du test")
        popup.geometry("350x350")
        popup.grab_set()

        # Déballage plus clair
        test_obj, test_type, *rest = test_info
        col1 = rest[0] if len(rest) > 0 else None
        col2 = rest[1] if len(rest) > 1 else None
        val1 = rest[2] if len(rest) > 2 else None
        val2 = rest[3] if len(rest) > 3 else None

        # Titre et type
        tk.Label(popup, text=f"Nom : {test_obj.nom}", font=("Segoe UI", 10, "bold")).pack(pady=5)
        tk.Label(popup, text=f"Type : {test_type}", font=("Segoe UI", 10, "italic")).pack(pady=5)

        # Fonction pour créer la liste des champs
        def creer_champs(test_obj, test_type, col1, col2, val1, val2):
            if isinstance(test_obj, Test_spe):
                champs_spe = {
                    "type_test": [("Type de test", "Test spécifique")],
                    "val_min": [("Colonne", col1), ("Valeur Min", val1)],
                    "val_max": [("Colonne", col1), ("Valeur Max", val1)],
                    "val_entre": [("Colonne", col1), ("Valeur Min", val1),  ("Valeur Max", val2)],
                    "compare_fix": [("Colonne 1", col1), ("Colonne 2", col2), ("Différence max", val1)],
                    "compare_ratio": [("Colonne 1", col1), ("Colonne 2", col2), ("Ratio autorisé", val1)],
                }
                return champs_spe.get("type_test", []) + champs_spe.get(test_type, [])
            elif isinstance(test_obj, Test_gen):
                critere_str = ", ".join(test_obj.critere)
                champs_gen = {
                    "type_test": [("Type de test", "Test générique")],
                    "val_min": [("Critères", critere_str), ("Valeur Min", col1 if col1 is not None else "N/A")],
                    "val_max": [("Critères", critere_str), ("Valeur Max", col2 if col2 is not None else "N/A")],
                    "val_entre": [("Critères", critere_str), ("Valeur Min", val1 if val1 is not None else "N/A"), ("Valeur Max", val2 if val2 is not None else "N/A")],
                }
                return champs_gen.get("type_test", []) + champs_gen.get(test_type, [])
            else:
                return [("Type", "Inconnu")]

        # Création et affichage
        champs = creer_champs(test_obj, test_type, col1, col2, val1, val2)

        for champ, valeur in champs:
            if valeur is not None:
                tk.Label(popup, text=f"{champ} : {valeur}").pack(anchor="w", padx=20)



# Affichage de l'aperçu du fichier Excel
    def create_excel_preview_frame(self):
        """Crée le cadre pour l'aperçu du fichier Excel."""
        # Créer un LabelFrame
        self.excel_preview_frame = tk.LabelFrame(self.scrollable_frame, text="3. Aperçu du fichier Excel", bg="#f4f4f4")
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


# Champs de résultats
    def create_results_frame(self):
        """Crée le cadre pour afficher les résultats des tests."""
        # Cadre pour les résultats du test
        self.results_frame = tk.LabelFrame(self, text="4. Résultats du test", bg="#f4f4f4")
        self.results_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Barre de défilement verticale
        scroll_y = tk.Scrollbar(self.results_frame, orient="vertical")
        scroll_y.pack(side="right", fill="y")

        # # Barre de défilement horizontale
        scroll_x = tk.Scrollbar(self.results_frame, orient="horizontal")
        scroll_x.pack(side="bottom", fill="x")

        # Zone de texte avec padding
            # Zone de texte avec padding
        self.result_text = tk.Text(
            self.results_frame, 
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





        return self.results_frame

# Affichage des détails des erreurs 
    def create_error_details_frame(self):
        """Crée le cadre pour afficher les détails des erreurs."""
        self.error_details_frame = tk.LabelFrame(self.scrollable_frame, text="5. Détails des erreurs", bg="#f4f4f4")
        self.error_details_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Frame pour contenir la Treeview et les scrollbar
        tree_frame = tk.Frame(self.error_details_frame, bg="#f4f4f4")
        tree_frame.pack(fill="both", expand=True)

        # La Treeview
        self.erreur_table = ttk.Treeview(tree_frame, columns=("Ligne", "Colonne", "Valeur"), show="headings")
        self.erreur_table.heading("Ligne", text="Ligne")
        self.erreur_table.heading("Colonne", text="Colonne")
        self.erreur_table.heading("Valeur", text="Valeur d'erreur")
        self.erreur_table.pack(side="left", fill="both", expand=True)

        # Barre de défilement verticale
        err_scroll_y = tk.Scrollbar(tree_frame, orient="vertical", command=self.erreur_table.yview)
        err_scroll_y.pack(side="right", fill="y")

        # Barre de défilement horizontale
        err_scroll_x = tk.Scrollbar(self.error_details_frame, orient="horizontal", command=self.erreur_table.xview)
        err_scroll_x.pack(fill="x")

        # Lier les scrollbar à la Treeview
        self.erreur_table.configure(yscrollcommand=err_scroll_y.set, xscrollcommand=err_scroll_x.set)

        return self.error_details_frame
        # texte.config(state="disabled")



# Fonctionnalités d'initialisation et de préparation des dossiers =========================================================================================================
    def prepare_dossiers(self):
        """Crée les dossiers nécessaires pour l'application."""
        Path("sauvegardes/sauvegardes_tests").mkdir(exist_ok=True)
        Path("sauvegardes/results").mkdir(exist_ok=True)
        Path("sauvegardes/data").mkdir(exist_ok=True)
            


# Verification d'une entrée entière =========================================================================================================
    def validate_integer_input(self, P):
        """
        Valide si l'entrée est un entier positif ou vide.
        """
        return P == "" or P.isdigit()

    def on_key_release_int(self, event):
        """Valide l'entrée de la taille de l'en-tête pour s'assurer qu'elle est un entier positif."""
        if not self.taille_entete_entry.get().isdigit() and self.taille_entete_entry.get() != "":
            messagebox.showwarning("Validation", "Veuillez entrer un nombre entier.")
            self.taille_entete_entry.delete(0, tk.END)


# Construction de la structure de l'entête =========================================================================================================
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

            return self.dico_structure

        except Exception as e:
            messagebox.showerror("Erreur", "Fichier et taille d'entete requis.")
            # messagebox.showerror("Erreur", f"Impossible de construire le dictionnaire d'en-tête : {e}")
            return {}
   

# Activation/desactivation des element =========================================================================================================
    def activation_bouton(self):
        self.taille_entete_entry.config(state="normal")
        self.detail_btn.config(state="normal")

        self.btn_popup_ajouter_test_gen.config(state="normal")
        self.btn_popup_ajouter_test_spe.config(state="normal")
        self.btn_executer_tests.config(state="normal")
        self.btn_sauvegarder_tests.config(state="normal")
        self.btn_importer_tests.config(state="normal")

    def desactivation_bouton(self):
        #entete
        self.taille_entete_entry.config(state="disabled")
        self.detail_btn.config(state="disabled")

        self.btn_popup_ajouter_test_gen.config(state="disabled")
        self.btn_popup_ajouter_test_spe.config(state="disabled")
        self.btn_executer_tests.config(state="disabled")
        self.btn_sauvegarder_tests.config(state="disabled")
        self.btn_importer_tests.config(state="disabled")




# Definition des tests =========================================================================================================

    def popup_ajouter_test_gen(self):
        """Ouvre un popup pour ajouter un test générique."""
        popup = tk.Toplevel(self)
        popup.title("Ajouter un test générique")
        popup.grab_set()

        # Nom du test
        tk.Label(popup, text="Nom du test :").grid(row=0, column=0, sticky="w")
        nom_entry = tk.Entry(popup, width=30)
        nom_entry.grid(row=0, column=1)

        # Symbole ou valeur à rechercher
        tk.Label(popup, text="Valeur/Symbole à rechercher :").grid(row=1, column=0, sticky="w")
        valeur_entry = tk.Entry(popup, width=30)
        valeur_entry.grid(row=1, column=1)

        # Type de test
        tk.Label(popup, text="Type de test :").grid(row=2, column=0, sticky="w")
        type_test = ttk.Combobox(popup, values=["val_min", "val_max", "val_entre"], state="readonly")
        type_test.grid(row=2, column=1)
        type_test.set("val_min")

        # Champs dynamiques
        label_val_min = tk.Label(popup, text="Valeur minimale :")
        val_min_entry = tk.Entry(popup)

        label_val_max = tk.Label(popup, text="Valeur maximale :")
        val_max_entry = tk.Entry(popup)

        # Masquer tous les champs dynamiques au départ
        for widget in [label_val_min, val_min_entry, label_val_max, val_max_entry]:
            widget.grid_forget()

        def afficher_champs_selon_type(event=None):
            """Affiche uniquement les champs nécessaires en fonction du type de test."""
            for widget in [label_val_min, val_min_entry, label_val_max, val_max_entry]:
                widget.grid_forget()

            t = type_test.get()
            ligne = 3  # Ligne de départ pour les champs dynamiques

            if t == "val_min":
                label_val_min.grid(row=ligne, column=0, sticky="w")
                val_min_entry.grid(row=ligne, column=1)
            elif t == "val_max":
                label_val_max.grid(row=ligne, column=0, sticky="w")
                val_max_entry.grid(row=ligne, column=1)
            elif t == "val_entre":
                label_val_min.grid(row=ligne, column=0, sticky="w")
                val_min_entry.grid(row=ligne, column=1)
                ligne += 1
                label_val_max.grid(row=ligne, column=0, sticky="w")
                val_max_entry.grid(row=ligne, column=1)

        # Lier l'événement de changement de type de test
        type_test.bind("<<ComboboxSelected>>", afficher_champs_selon_type)
        afficher_champs_selon_type()  # Afficher les champs initiaux

        def ajouter():
            """Ajoute le test générique."""
            nom = nom_entry.get().strip()
            valeur = valeur_entry.get().strip()
            type_selected = type_test.get()

            try:
                val_min = float(val_min_entry.get()) if val_min_entry.winfo_ismapped() and val_min_entry.get() else None
                val_max = float(val_max_entry.get()) if val_max_entry.winfo_ismapped() and val_max_entry.get() else None
            except ValueError:
                messagebox.showerror("Erreur", "Veuillez entrer des valeurs numériques valides.")
                return

            if not nom or not valeur:
                messagebox.showerror("Erreur", "Le nom du test et la valeur/symbole sont requis.")
                return

            test = Test_gen(nom=nom, critere=[valeur])
            self.tests.append((test, type_selected, val_min, val_max))
            self.test_listbox.insert(tk.END, f"[GEN] {nom} ({type_selected})")
            popup.destroy()

        # Bouton pour ajouter le test
        ttkb.Button(popup, text="Ajouter le test", command=ajouter).grid(row=6, column=1, pady=10)

    def popup_ajouter_test_spe(self):
        """Ouvre un popup pour ajouter un test spécifique."""
        if self.taille_entete_entry.get() == "":
                messagebox.showerror("Erreur", "Veuillez entrer la taille de l'en-tête.")
                return 
        try:
            
            dico = self.dico_entete()  # Assure que self.dico_structure est construit
        except Exception as e:
            messagebox.showerror("Erreur", "Fichier et taille d'entete requis.")
            return
        if dico == {}:
            return

        popup = tk.Toplevel(self)
        popup.title("Ajouter un test spécifique")
        popup.grab_set()

        ligne  = 0

        # Nom du test
        tk.Label(popup, text="Nom du test :").grid(row=ligne, column=0, sticky="w")
        nom_entry = tk.Entry(popup, width=30)
        nom_entry.grid(row=ligne, column=1)

        ligne+=1
        # Type de test
        tk.Label(popup, text="Type de test :").grid(row=ligne, column=0, sticky="w")
        type_test = ttk.Combobox(popup, values=["val_min", "val_max", "val_entre", "compare_fix", "compare_ratio"], state="readonly")
        type_test.grid(row=ligne, column=1)
        type_test.set("val_min")

        ligne+=1
        label_col1 = tk.Label(popup, text="Colonne cible 1 :")
        label_col1.grid(row=ligne, column=0, sticky="w")
        colonne_cible_1_combo = Selection_col(dico)
        colonne_cible_1_combo.get_frame_selection_grid(popup,ligne,1)

        ligne+= int(self.taille_entete_entry.get())

        label_col2 = tk.Label(popup, text="Colonne cible 2 :")
        label_col2.grid(row=ligne, column=0, sticky="w")
        colonne_cible_2_combo = Selection_col(dico)
        colonne_cible_2_combo.get_frame_selection_grid(popup,ligne,1)

        
        ligne+=int(self.taille_entete_entry.get())


        # Champs dynamiques selon le type de test

        label_val_min = tk.Label(popup, text="Valeur minimale :")
        val_min_entry = tk.Entry(popup)

        label_val_max = tk.Label(popup, text="Valeur maximale :")
        val_max_entry = tk.Entry(popup)

        label_diff = tk.Label(popup, text="Différence attendue :")
        diff_entry = tk.Entry(popup)

        label_ratio = tk.Label(popup, text="Ratio attendu :")
        ratio_entry = tk.Entry(popup)

        for widget in [label_val_min, val_min_entry, label_val_max, val_max_entry, label_diff, diff_entry, label_ratio, ratio_entry, label_col2]:
            widget.grid_forget()
        colonne_cible_2_combo.grid_remove()

        def afficher_champs_selon_type(event=None,ligne=5):
            for widget in [label_val_min, val_min_entry, label_val_max, val_max_entry, label_diff, diff_entry, label_ratio, ratio_entry,  label_col2]:
                widget.grid_forget()
            colonne_cible_2_combo.grid_remove()

            t = type_test.get()
            ligne_i = ligne
            print(ligne_i)
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

            elif t =="compare_fix"or t == "compare_ratio":
                print(ligne_i)
                label_col2.grid(row=ligne_i, column=0, sticky="w")
                colonne_cible_2_combo.grid()
                ligne_i += int(self.taille_entete_entry.get())

                # label_val_min.grid(row=ligne_i, column=0, sticky="w")
                # val_min_entry.grid(row=ligne_i, column=1)
                ligne_i +=1

                if t == "compare_fix":
                    label_diff.grid(row=ligne_i, column=0, sticky="w")
                    diff_entry.grid(row=ligne_i, column=1)
                    ligne_i +=1

                else:
                    label_ratio.grid(row=ligne_i, column=0, sticky="w")
                    ratio_entry.grid(row=ligne_i, column=1)
                    ligne_i +=1



        type_test.bind("<<ComboboxSelected>>", afficher_champs_selon_type)
        afficher_champs_selon_type(ligne=ligne)

        def ajouter():
            nom = nom_entry.get().strip()
            type_selected = type_test.get()


            # Chemins complets des colonnes
            chemin_1 = colonne_cible_1_combo.chemin
            chemin_2 = colonne_cible_2_combo.chemin

            try:
                if type_selected == "val_min" or type_selected == "val_entre":
                    val1 = float(val_min_entry.get()) if val_min_entry.get() else None
                elif type_selected =="compare_fix":
                    val1 = float(diff_entry.get()) if diff_entry.get() else None
                elif type_selected == "compare_ratio":
                    val1 = float(ratio_entry.get()) if ratio_entry.get() else None
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

        ttkb.Button(popup, text="Ajouter le test", command=ajouter).grid(row=ligne + 6, column=1, pady=10)

# Exécuter les tests =========================================================================================================

    def executer_tests(self):
        """Exécute les tests sélectionnés sur le fichier Excel."""
        if self.taille_entete_entry.get() == "":
            messagebox.showerror("Erreur", "Veuillez entrer la taille de l'en-tête.")
            return 
        if self.tests == []:
            messagebox.showwarning("Attention", "Aucun test sélectionné.")
            return
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


        for item in self.erreur_table.get_children():
            self.erreur_table.delete(item)

        for test in self.tests:
            message = ""
            if isinstance(test[0], Test_gen):
                obj, type_test, val_min, val_max = test
                self.append_text(f"--- {obj.nom} ({type_test}) ---")
                try:
                    if type_test == "val_min":
                        message = obj.val_min(feuille, val_min)
                    elif type_test == "val_max":
                        message = obj.val_max(feuille, val_max)
                    elif type_test == "val_entre":
                        message = obj.val_entre(feuille, val_min, val_max)
                    
                    self.append_text(str(message))

                    

                except Exception as e:
                    self.append_text(f"Erreur test {obj.nom}: {e}", color="red")


            elif isinstance(test[0], Test_spe):
                # ⬇️ décomposition étendue avec les nouvelles cases à cocher
                obj, type_test, col1, col2, val1, val2 = test
                obj.feuille = feuille  # mise à jour de la feuille
            
                self.append_text(f"--- {obj.nom} ({type_test}) ---")
                
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
            
                    self.append_text(str(message))
            
                except Exception as e:
                    self.append_text(f"Erreur test {obj.nom}: {e}", color="red")



            self.append_text(f"--- FIN TESTS ---\n")
        feuille.error_all_cell_colors()
        for row_idx, ligne in enumerate(feuille.erreurs):
            for col_idx, code in enumerate(ligne):
                if code > 0:
                    self.erreur_table.insert("", "end", values=(row_idx + 1, col_idx + 1, feuille.df.iloc[row_idx, col_idx]))





