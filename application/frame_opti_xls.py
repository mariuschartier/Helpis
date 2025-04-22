import pandas as pd
import opti_fichier  # ton module de conversion
from tkinter import filedialog, messagebox, ttk
import tkinter as tk
from pathlib import Path


class opti_xls(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        
        self.fichier_path = None
        self.df = None
        self.feuille_nom = tk.StringVar()
        
        self.prepare_dossiers()
        
        frame_fichier = tk.LabelFrame(self, text="1. Charger un fichier Excel", bg="#f4f4f4")
        frame_fichier.pack(fill="x", padx=10, pady=5)

        self.fichier_entry = tk.Entry(frame_fichier, width=80)
        self.fichier_entry.pack(side="left", padx=5, pady=5)

        tk.Button(frame_fichier, text="Parcourir fichier .xls", command=self.controller.bind_button(self.choisir_fichier)).pack(side="left", padx=5)
        tk.Button(frame_fichier, text="Convertir en .xlsx", command=controller.bind_button(self.convertir_fichier)).pack(side="left", padx=5)

        self.status_label = tk.Label(self, text="", bg="#f4f4f4", fg="green")
        self.status_label.pack(pady=5)
        
        
        
    def prepare_dossiers(self):
        Path("sauvegardes_tests").mkdir(exist_ok=True)
        Path("results").mkdir(exist_ok=True)
        Path("data").mkdir(exist_ok=True)
        
    def choisir_fichier(self):
        dossier_data = Path("data")
        dossier_data.mkdir(parents=True, exist_ok=True)  # Crée le dossier s’il n’existe pas
    
        filepath = filedialog.askopenfilename(
            filetypes=[("Fichiers Excel 97-2003", "*.xls")],
            initialdir=dossier_data,  # Dossier par défaut
            title="Choisir un fichier .xls"
        )
        if filepath:
            self.fichier_path = filepath
            self.fichier_entry.delete(0, tk.END)
            self.fichier_entry.insert(0, filepath)
    
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
    
