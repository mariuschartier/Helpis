

import tkinter as tk
from tkinter import messagebox

class MonApplication:
    def __init__(self, root):
        self.root = root
        self.root.title("Ma première application Tkinter")

        # Créer un label
        self.label = tk.Label(root, text="Entrez votre nom :")
        self.label.pack(pady=10)

        # Champ de texte
        self.entry = tk.Entry(root)
        self.entry.pack(pady=5)

        # Bouton
        self.bouton = tk.Button(root, text="Valider", command=self.dire_bonjour)
        self.bouton.pack(pady=10)

    def dire_bonjour(self):
        nom = self.entry.get()
        messagebox.showinfo("Bonjour", f"Bonjour, {nom} !")

if __name__ == "__main__":
    root = tk.Tk()
    app = MonApplication(root)
    root.mainloop()