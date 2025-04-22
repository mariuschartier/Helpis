import tkinter as tk
from tkinter import ttk
from app import ExcelTesterApp
from frame_opti_xls import opti_xls

class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Application Multi-pages")
        self.geometry("400x300")
        
        # Créer un conteneur pour les pages
        self.frames = {}
        
        for F in (opti_xls, ExcelTesterApp):
            frame = F(parent=self)
            self.frames[F] = frame
            frame.grid(row=0, column=0, sticky="nsew")
        
        self.create_navigation()
        self.show_frame(opti_xls)

    def create_navigation(self):
        nav_frame = tk.Frame(self)
        nav_frame.grid(row=1, column=0, sticky="ew")  # Utilisez grid ici au lieu de pack

        home_button = tk.Button(nav_frame, text="Accueil", command=lambda: self.show_frame(opti_xls))
        home_button.grid(row=0, column=0, padx=5, pady=5)  # Utiliser grid au lieu de pack
        
        info_button = tk.Button(nav_frame, text="Informations", command=lambda: self.show_frame(ExcelTesterApp))
        info_button.grid(row=0, column=1, padx=5, pady=5)  # Utiliser grid au lieu de pack

    def show_frame(self, page):
        frame = self.frames[page]
        frame.tkraise()

class HomePage(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        label = tk.Label(self, text="Bienvenue sur la page d'accueil", font=("Helvetica", 16))
        label.pack(pady=20)

class InfoPage(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        label = tk.Label(self, text="Voici des informations supplémentaires.", font=("Helvetica", 16))
        label.pack(pady=20)

if __name__ == "__main__":
    app = Application()
    app.mainloop()