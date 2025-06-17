import tkinter as tk
from tkinter import ttk
from tkinter import messagebox

def action_cellule(event):
    # Identifier la région cliquée
    region = tree.identify('region', event.x, event.y)
    if region == 'cell':
        # Identifier la ligne et la colonne
        row_id = tree.identify_row(event.y)
        col_id = tree.identify_column(event.x)
        # Vérifier si c'est la colonne 'Action'
        if col_id == '#2':  # La deuxième colonne
            item = tree.item(row_id)
            values = item['values']
            # Appeler la fonction spécifique
            messagebox.showinfo("Action", f"Action sur {values[0]}")
        # Ajoutez d'autres conditions si nécessaire

def on_heading_click(event):
    region = tree.identify('region', event.x, event.y)
    if region == 'heading':
        col = tree.identify_column(event.x)
        if col == '#2':  # La colonne "Action"
            messagebox.showinfo("En-tête", "Vous avez cliqué sur l'en-tête de 'Action'!")

root = tk.Tk()
root.title("Exemple avec clics sur cellules et en-tête")

# Création du Treeview
tree = ttk.Treeview(root, columns=("Nom", "Action"), show='headings')

# Définir les en-têtes
tree.heading("Nom", text="Nom")
tree.heading("Action", text="Action")

# Définir la largeur des colonnes
tree.column("Nom", width=150)
tree.column("Action", width=150)

# Ajouter des données
for i in range(10):
    tree.insert("", "end", values=(f"Item {i+1}", "Cliquez ici"))

tree.pack(fill='both', expand=True)

# Lier le clic pour la cellule
tree.bind("<Button-1>", action_cellule)
# Lier le clic pour l'en-tête
tree.bind("<Button-1>", on_heading_click, add='+')

root.mainloop()