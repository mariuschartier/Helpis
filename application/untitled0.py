import tkinter as tk
from tkinter import ttk
from app import ExcelTesterApp
from frame_opti_xls import opti_xls




class MultiPageApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Application Multi-page - Excel Tool")
        self.geometry("1000x700")
        self.configure(bg="#f4f4f4")

        # Apparence
        style = ttk.Style()
        style.theme_use('clam')
        font_default = ('Segoe UI', 10)

        style.configure(".", font=font_default)
        style.configure("TButton", font=font_default, padding=6, relief="flat",
                        background="#005f73", foreground="white")
        style.map("TButton", background=[('active', '#0a9396'), ('disabled', '#cccccc')])
        style.configure("Treeview", font=font_default, rowheight=24,
                        background='white', foreground='black', fieldbackground='white')
        style.configure("Treeview.Heading", font=('Segoe UI', 10, 'bold'),
                        background="#94d2bd", foreground="black")
        style.configure("TLabel", background="#f4f4f4", font=font_default)

        # Navigation
        nav_frame = tk.Frame(self, bg="#e0e0e0", pady=5)
        nav_frame.pack(side="top", fill="x")
        tk.Button(nav_frame, text="📊 Tests Excel", command=lambda: self.afficher_page("tests")).pack(side="left", padx=5)
        tk.Button(nav_frame, text="📁 Conversion XLS → XLSX", command=lambda: self.afficher_page("convert")).pack(side="left", padx=5)
        tk.Button(nav_frame, text="❓ Aide", command=self.ouvrir_aide).pack(side="right", padx=5)

        # Container
        self.container = tk.Frame(self, bg="#f4f4f4")
        self.container.pack(fill="both", expand=True)

        self.pages = {}
        self.init_pages()
        self.afficher_page("tests")

    def init_pages(self):
        self.pages["tests"] = ExcelTesterApp(self.container)
        self.pages["convert"] = opti_xls(self.container)

    def afficher_page(self, nom_page):
        for page in self.pages.values():
            page.pack_forget()
        if nom_page in self.pages:
            self.pages[nom_page].pack(fill="both", expand=True)
        else:
            print(f"Erreur : page '{nom_page}' non trouvée.")

    def ouvrir_aide(self):
        aide_popup = tk.Toplevel(self)
        aide_popup.title("Aide - Utilisation")
        aide_popup.geometry("600x400")
        texte = tk.Text(aide_popup, wrap="word", font=("Segoe UI", 10))
        texte.pack(fill="both", expand=True, padx=10, pady=10)
        contenu = (
            "Bienvenue dans l'application Excel multi-fonctions 🧪\n"
            "1. Tests Excel : permet d'ajouter des règles et de colorer les cellules erronées.\n"
            "2. Conversion : permet de convertir les fichiers .xls vers .xlsx.\n"
            "3. Résultats : erreurs détectées affichées et enregistrées dans le fichier.\n"
        )
        texte.insert(tk.END, contenu)
        texte.config(state="disabled")

if __name__ == "__main__":
    app = MultiPageApp()
    app.mainloop()
