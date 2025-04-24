import tkinter as tk
from tkinter import ttk
from app import ExcelTesterApp
from frame_opti_xls import opti_xls
import threading



class MultiPageApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Application Multi-page - Excel Tool")
        self.geometry("1000x700")
        self.configure(bg="#f4f4f4")
        print(self.winfo_width())
        # Apparence
        # 🎨 Couleurs harmonisées
        BLEU_PROFOND = "#005f73"
        VERT_EAU = "#0a9396"
        VERT_CLAIR = "#94d2bd"
        FOND_CLAIR = "#e0fbfc"
        TEXTE_BLEU = "#0077b6"
        TEXTE_FOND = "#f4f4f4"
        
        # Apparence
        style = ttk.Style()
        style.theme_use('clam')
        font_default = ('Segoe UI', 10)
        
        # Style global
        style.configure(".", font=font_default, background=FOND_CLAIR)
        
        # Boutons
        style.configure("TButton",
            font=font_default,
            padding=6,
            relief="flat",
            background=BLEU_PROFOND,
            foreground="white")
        style.map("TButton",
            background=[('active', VERT_EAU), ('disabled', '#cccccc')],
            foreground=[('pressed', 'white'), ('active', 'white')])
        
        # Tableaux
        style.configure("Treeview",
            font=font_default,
            rowheight=24,
            background="white",
            fieldbackground="white",
            foreground="black")
        style.configure("Treeview.Heading",
            font=('Segoe UI', 10, 'bold'),
            background=VERT_CLAIR,
            foreground="black")
        
        # Labels
        style.configure("TLabel", background=FOND_CLAIR, font=font_default)
        self.configure(bg=FOND_CLAIR)
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
        
        # icon de cchargement
        self.loading_label = tk.Label(self.container, text="", fg="green", font=("Segoe UI", 10, "italic"))
        self.loading_label.pack(pady=5)
        self.hide_loading()

    def init_pages(self):
        self.pages["tests"] = ExcelTesterApp(self.container, controller=self)
        self.pages["convert"] = opti_xls(self.container, controller=self)

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
        
    def show_loading(self, message="⏳ Chargement..."):
        if hasattr(self, 'loading_label'):
            self.loading_label.config(text=message)
        else:
            self.loading_label = tk.Label(self, text=message, bg="#f4f4f4", fg="blue", font=("Segoe UI", 10, "italic"))
            self.loading_label.pack(side="top", fill="x", pady=2)
    
    def hide_loading(self):
        if hasattr(self, 'loading_label'):
            self.loading_label.destroy()
            del self.loading_label
        


    def exec_with_loading(self, func):
        # Afficher le label de chargement
        self.loading_label = tk.Label(
            self.container,
            text="⏳ Chargement en cours, veuillez patienter...",
            font=("Segoe UI", 11, "italic"),
            fg="#0066cc",     # bleu doux
            bg="#f4f4f4",     # fond cohérent
        )
        self.loading_label.pack(pady=10)
                 
        def task():
            try:
                func()
            except Exception as e:
                print("Erreur pendant l'exécution :", e)
            finally:
                # Revenir dans le thread principal pour enlever le label
                self.after(0, self.loading_label.destroy)
    
        threading.Thread(target=task, daemon=True).start()


    def bind_button(self, action):
        """Associe une fonction à exécuter via le contrôleur avec animation de chargement."""
        return lambda: self.exec_with_loading(action)
    
    
 
        


if __name__ == "__main__":
    app = MultiPageApp()
    app.mainloop()
