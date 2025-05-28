import tkinter as tk
from tkinter import colorchooser

# Fonction pour changer la couleur de fond du bouton
def changer_couleur():
    # Ouvre une boîte de dialogue pour choisir une couleur
    couleur = colorchooser.askcolor(title="Choisissez une couleur")
    if couleur[1]:  # Si une couleur a été sélectionnée
        bouton.config(bg=couleur[1])

# Créer la fenêtre principale
fenetre = tk.Tk()
fenetre.title("Manipulation des couleurs avec Tkinter")
fenetre.geometry("300x200")

# Créer un bouton
bouton = tk.Button(fenetre, text="Changer la couleur de fond", command=changer_couleur, bg="lightgray")
bouton.pack(pady=50)

# Lancer la boucle principale
fenetre.mainloop()