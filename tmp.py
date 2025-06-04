import ttkbootstrap as ttkb
from tkinter import ttk

app = ttkb.Window(themename='litera')

style = ttkb.Style()

# Définir ton style pour la frame
style.configure(
    "Custom.TFrame",
    background="#ffffff",  # Couleur de fond
    borderwidth=0,
    relief='raised'
)

# Créer la frame avec le style
frame = ttkb.Frame(app, style="Custom.TFrame")
frame.pack(fill='both', expand=True, padx=20, pady=20)

# Ajoute un contenu dans la frame, par exemple un label
label = ttkb.Label(frame, text="Contenu de la frame")
label.pack(pady=10)

# Ajouter une ligne en bas avec un Separator
separator = ttk.Separator(app, orient='horizontal')
separator.pack(fill='x', padx=20, pady=(0, 10))  # Paddings pour positionner la ligne

app.mainloop()
