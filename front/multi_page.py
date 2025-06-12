import ttkbootstrap as ttkb
from ttkbootstrap.constants import *
import threading
from tkinter import filedialog, messagebox, ttk

from front.frame_opti_xls import opti_xls
from front.page_comparaison import ComparePage
from front.app import ExcelTesterApp

class MultiPageApp(ttkb.Window):
    """Application multi-page pour la manipulation de fichiers Excel."""
    def __init__(self):
        super().__init__(themename="flatly")  # ou autre thème
        self.title("Application Multi-page - Excel Tool")
        self.state('zoomed')
        self.iconbitmap('logo.ico')
        # Couleurs et style
        self.configure(bg="#f4f4f4")
        style = ttkb.Style()
        style.configure(".", font=('Segoe UI', 10), background="#f4f4f4")
        
        # Style spécifique pour les boutons
        self.setup_styles()
        # Barre de navigation
        nav_frame = ttkb.Frame(self, style="Custom.TFrame")
        nav_frame.pack(side="top", fill="x")
        ttkb.Button(nav_frame, text="📊 Verification Excel",  style = "Nav.TButton", command=lambda: self.afficher_page("tests")).pack(side="left", padx=5)
        ttkb.Button(nav_frame, text="📁 Manipulation",  style = "Nav.TButton", command=lambda: self.afficher_page("convert")).pack(side="left", padx=5)
        ttkb.Button(nav_frame, text="📈 Tests Statistiques",  style = "Nav.TButton", command=lambda: self.afficher_page("compare")).pack(side="left", padx=5)
        ttkb.Button(nav_frame, text="❓ Aide",  style = "Nav.TButton", command=self.ouvrir_aide).pack(side="right", padx=5)

        # Conteneur pour les pages
        self.container = ttkb.Frame(self, style="Custom.TFrame")
        self.container.pack(fill="both", expand=True)

        self.pages = {}
        self.init_pages()
        self.afficher_page("tests")
        
        # Label de chargement
        self.loading_label = ttkb.Label(self.container, text="", style="TLabel")
        self.loading_label.pack(pady=5)
        self.hide_loading()

    def init_pages(self):
        self.pages["tests"] = ExcelTesterApp(self.container, controller=self)
        self.pages["convert"] = opti_xls(self.container, controller=self)
        self.pages["compare"] = ComparePage(self.container, controller=self)

    def afficher_page(self, nom_page):
        for page in self.pages.values():
            page.pack_forget()
        if nom_page in self.pages:
            self.pages[nom_page].pack(fill="both", expand=True)
        else:
            print(f"Erreur : page '{nom_page}' non trouvée.")

    def ouvrir_aide(self):
        aide_popup = ttkb.Toplevel(self)
        aide_popup.title("Aide - Utilisation")
        aide_popup.geometry("600x400")
        texte = ttkb.Text(aide_popup, wrap="word", font=('Segoe UI', 10))
        texte.pack(fill="both", expand=True, padx=10, pady=10)
        contenu = (
            "Bienvenue dans l'application Excel multi-fonctions 🧪\n"
            "1. 📊 Verification Excel : permet d'effectuer des tests pour détecter des erreurs dans les fichiers .xlsx.\n"
            "2. 📁 Manipulation : permet de manipuler les fichiers .xls et .xlsx(conversion de .xls à .xlsx, formatage de .xlsx).\n"
            "3. 📈 Tests Statistiques : permet d'effectuer des tests statistiques sur les fichiers .xlsx\n"
        )
        texte.insert("end", contenu)
        texte.configure(state="disabled")
        
    def show_loading(self, message="⏳ Chargement..."):
        if hasattr(self, 'loading_label'):
            self.loading_label.config(text=message)
        else:
            self.loading_label = ttkb.Label(self, text=message, style="TLabel")
            self.loading_label.pack(side="top", fill="x", pady=2)
    
    def hide_loading(self):
        if hasattr(self, 'loading_label'):
            self.loading_label.destroy()
            del self.loading_label
    
    def exec_with_loading(self, func):
        self.show_loading()
        def task():
            try:
                func()
            except Exception as e:
                print("Erreur pendant l'exécution :", e)
                messagebox.showerror("Erreur", f"Erreur lors de la creation : {e}")
            finally:
                self.after(0, self.hide_loading)
        threading.Thread(target=task, daemon=True).start()

    def bind_button(self, action):
        return lambda: self.exec_with_loading(action)
    
    def setup_styles(self):
        style = ttkb.Style()
        
        
        # Palette de couleurs
        primary_color = "#ACACAC"    # Couleur principale (bleu foncé)
        accent_color = "#3399ff"     # Couleur d'accent (turquoise clair)
        background_color = "#ffffff" # Couleur de fond (gris clair)
        header_bg = "#cccccc"        # Couleur pour les en-têtes
        hover_color = "#090a0a"      # Couleur au survol (active)
        pressed_color = "#005f73"    # Couleur lors du clic

        # Police principale
        font_main = ('Segoe UI', 10)
        font_bold = ('Segoe UI', 10, 'bold')
        

        # Style général pour la fenêtre
        style.configure('.', background=background_color, font=font_main)

        # Style pour les Labels
        style.configure("TLabel", background=background_color, font=font_main)

        # Style pour les boutons
        style.configure("TButton",
                    font=font_main,
                    padding=6,
                    relief="flat",
                    background="#007BFF",
                    foreground="#ffffff")
        style.map("TButton",
                background=[('active', '#0056b3'), ('disabled', "#ACACAC")],
                foreground=[('pressed', '#ffffff'), ('active', '#ffffff'), ('active', '#ffffff')])
        # Style pour les LabelsFrame
        style.configure("TLabelframe",
                        background=background_color,
                        borderwidth=1,
                        relief="flat")
        style.configure("TLabelframe.Label",
                        background=background_color,
                        font=font_bold)
        style.configure(
            "Nav.TButton",
            font=('Helvetica', 12),
            padding=10,
            foreground='white',        # Couleur du texte
            background='#007BFF',      # Couleur de fond
     
            relief='raised'
        )

        # Modifier le style avec map pour changer l’apparence au survol ou lors du clic
        style.map(
            "Nav.TButton",
            background=[('active', '#0056b3')],
            relief=[('pressed', 'sunken')]
)




        # Style pour Treeview
        style.configure("Treeview",
                        font=font_main,
                        rowheight=24,
                        background="#ffffff",
                        fieldbackground="#ffffff",
                        foreground="#000000")
        style.configure("Treeview.Heading",
                        font=('Segoe UI', 10, 'bold'),
                        background=header_bg,
                        foreground="#000000")
        style.map("Treeview",
                background=[('selected', accent_color)],
                foreground=[('selected', 'white')])

        # Style pour scrollbar (optionnel, pour harmoniser)
        style.configure("Vertical.TScrollbar", gripcount=0,
                        background=primary_color,
                        troughcolor=background_color,
                        bordercolor=background_color,
                        arrowcolor="white")
        style.map("Vertical.TScrollbar",
                background=[('active', hover_color)])
  

        style.configure(
            "Custom.TFrame",
            background="#ffffff",  # Couleur de fond
            borderwidth=0,
            relief='raised',
        )

        style.configure(
            "NoBorder.TLabelframe",
            borderwidth=0,
            relief='flat'
        )


        style.configure(
            "Custom.TCombobox",
            foreground="#333333",  # Couleur du texte
            background="#e0e0e0",  # Couleur de fond
            selectbackground="#b0b0b0",  # Couleur lors de la sélection
            selectforeground="#000000",  # Couleur du texte sélectionné
            arrowcolor="#ff0000",  # Couleur de la flèche
            padding=5,
        )



    # Si vous utilisez d’autres widgets, vous pouvez continuer à personnaliser ici




# Lancement de l'application
if __name__ == "__main__":
    app = MultiPageApp()
    app.mainloop()