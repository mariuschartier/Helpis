from tests import opti_fichier  # ton module de conversion
from tkinter import filedialog, messagebox, ttk
import tkinter as tk
from pathlib import Path
from tests import opti_xlsx 
import pandas as pd

from structure.Fichier import Fichier
from structure.Feuille import Feuille


class opti_xls(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        
        self.fichier_path = None
        self.df = None
        self.feuille_nom = ""
        self.taille_entete_entry=0
        self.prepare_dossiers()
        
        self.choix_page(controller)
   
        
        frame_action = tk.LabelFrame(self, text="2. action sur le fichier Excel", bg="#f4f4f4")
        frame_action.pack(fill="x", padx=10, pady=5)
    
        tk.Button(frame_action, text="Convertir en .xlsx", command=controller.bind_button(self.convertir_fichier)).pack(side="left", padx=5)
        tk.Button(frame_action, text="ameliorer le .xlsx", command=controller.bind_button(self.ameliorer_fichier_xlsx)).pack(side="left", padx=5)
        tk.Button(frame_action, text="moyenne par jour", command=controller.bind_button(self.moyenne_par_jour)).pack(side="left", padx=5)
        tk.Button(frame_action, text="moyenne par semaine", command=controller.bind_button(self.moyenne_par_semaine)).pack(side="left", padx=5)


                

        self.status_label = tk.Label(self, text="", bg="#f4f4f4", fg="green")
        self.status_label.pack(pady=5)
        # self.create_excel_preview_frame()


    def choix_page(self,controller):
        self.frame_fichier = tk.LabelFrame(self, text="1. Charger un fichier Excel", bg="#f4f4f4")
        self.frame_fichier.pack(fill="x", padx=10, pady=5)

        self.fichier_entry = tk.Entry(self.frame_fichier, width=80)
        self.fichier_entry.pack(side="left", padx=5, pady=5)

        tk.Button(self.frame_fichier, text="Parcourir", command=self.controller.bind_button(self.choisir_fichier)).pack(side="left", padx=5)
        # Choix de la feuille
        self.feuille_combo = ttk.Combobox(self.frame_fichier, textvariable=self.feuille_nom, state="readonly")
        self.feuille_combo.pack(side="left", padx=5)
        self.feuille_combo.bind("<<ComboboxSelected>>", lambda e: self.afficher_excel())
    

        # Choix de la taille de l'en-tête
        tk.Label(self.frame_fichier, text="Taille de l'en-tête :").pack(side="left", padx=(10, 0))
        self.taille_entete_entry = tk.Entry(self.frame_fichier, width=5)
        self.taille_entete_entry.pack(side="left", padx=5)
         
        
    def prepare_dossiers(self):
        Path("sauvegardes_tests").mkdir(exist_ok=True)
        Path("results").mkdir(exist_ok=True)
        Path("data").mkdir(exist_ok=True)
        
    def choisir_fichier(self):
        dossier_data = Path("data")
        dossier_data.mkdir(parents=True, exist_ok=True)  # Crée le dossier s’il n’existe pas
    
        filepath = filedialog.askopenfilename(
            filetypes=[("Fichiers Excel 97-2003", "*.xls;*.xlsx")],
            initialdir=dossier_data,  # Dossier par défaut
            title="Choisir un fichier"
        )

        if filepath:
            self.fichier_path = filepath
            self.fichier_entry.delete(0, tk.END)
            self.fichier_entry.insert(0, filepath)

            try:
                xls = pd.ExcelFile(filepath)
                feuilles = xls.sheet_names
                self.feuille_combo['values'] = feuilles
                if feuilles:
                    self.feuille_combo.set(feuilles[0])
                    self.feuille_nom=feuilles[0]  # 👈 assure que StringVar est mis à jour
                    self.afficher_excel()
            except Exception as e:
                messagebox.showerror("Erreur", f"Impossible de lire les feuilles du fichier :\n{e}")
                
    
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
            opti_xlsx.process_excel_data(self.fichier_path,self.feuille_nom, fichier_destination)
            self.status_label.config(text="✅ Conversion terminée avec succès", fg="green")
            messagebox.showinfo("Succès", f"Fichier converti avec succès :\n{fichier_destination}")
        except Exception as e:
            self.status_label.config(text="❌ Échec de la conversion", fg="red")
            messagebox.showerror("Erreur", f"Erreur lors de la conversion : {e}")

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
            f1_1 = Feuille(f1,self.feuille_nom,int(self.taille_entete_entry.get()))
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
            f1_1 = Feuille(f1,self.feuille_nom,int(self.taille_entete_entry.get()))
            opti_xlsx.moyenne_par_semaine(f1_1,fichier_destination)
            self.status_label.config(text="✅ creation terminée avec succès", fg="green")
            messagebox.showinfo("Succès", f"Fichier creer avec succès :\n{fichier_destination}")
        except Exception as e:
            self.status_label.config(text="❌ Échec de la conversion", fg="red")
            messagebox.showerror("Erreur", f"Erreur lors de la creation : {e}")









    def create_excel_preview_frame(self):
        # Créer un LabelFrame
        self.excel_preview_frame = tk.LabelFrame(self, text="3. Aperçu du fichier Excel", bg="#f4f4f4")
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
    
    
    def afficher_excel(self):
        if not self.fichier_path or not self.feuille_nom:
            return
    
        try:
            # Vider l'ancien contenu
            for item in self.table.get_children():
                self.table.delete(item)
    
            # Lire le fichier Excel
            df = pd.read_excel(self.fichier_path, sheet_name=self.feuille_nom, header=None, engine="openpyxl")
            self.df = df
    
            # # Configurer les colonnes dans le Treeview
            # self.table["columns"] = list(range(len(df.columns)))
            # for i, col in enumerate(self.table["columns"]):
            #     self.table.heading(col, text=f"Col {col}")
            #     self.table.column(col, width=100)
    
            # # Ajouter les lignes (limité à 100 pour performance si besoin)
            # for i, row in df.iterrows():
            #     self.table.insert("", "end", values=list(row))
    
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d'afficher la feuille sélectionnée :\n{e}")

