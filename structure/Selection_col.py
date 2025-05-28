from tkinter import ttk
import tkinter as tk


class Selection_col:
    """
    Classe pour gérer la sélection de colonnes et sous-catégories dans une structure d'entête."""

    def __init__(self, structure_entete=None):
        """
        Initialise la classe avec une structure d'entête."""
        self.structure_entete = structure_entete or {}
        self.chemin = ""
        self.widgets = []  # pour suivre tous les widgets utilisés (combo inclus)
        self.action_selection = None


    def maj_donnees(self, structure_entete):
        """Met à jour la structure d'entête avec une nouvelle structure."""
        self.structure_entete = structure_entete
        self.colonne_combo['values'] = list(self.structure_entete.keys())
        # print(self.structure_entete)

    
    def get_frame_selection_grid(self, parent_frame, start_row=0, start_col=0):
        """
        Crée une interface en grid pour sélectionner une colonne et ses sous-catégories, et renvoie le chemin sélectionné.

        Args:
        - parent_frame (tk.Frame): Cadre parent où seront placés les widgets.
        - start_row (int): Ligne de départ pour le placement.

        Returns:
        - function: Fonction qui retourne le chemin sélectionné.
        """

        self.colonne_combo = ttk.Combobox(parent_frame, values=list(self.structure_entete.keys()), state="readonly")
        self.colonne_combo.grid(row=start_row, column=start_col, padx=5, pady=5, sticky="w")
        self.widgets.append(self.colonne_combo)

        comboboxes = []

        def add_combobox_grid(level, structure):
            combo = ttk.Combobox(parent_frame, state="readonly")
            combo.grid(row=start_row + level + 1, column=start_col, padx=5, pady=2, sticky="w")
            combo["values"] = list(structure.keys())
            comboboxes.append((combo, structure))
            self.widgets.append(combo)

            def on_selection(event=None):
                while len(comboboxes) > level + 1:
                    widget, _ = comboboxes.pop()
                    widget.destroy()

                selection = combo.get()
                if selection in structure and isinstance(structure[selection], dict) and structure[selection] and structure[selection] != {} :
                    print(f"Ajout de la combobox pour {selection}")
                    add_combobox_grid(level + 1, structure[selection])

                self.chemin = get_path()
                # print(self.chemin)
            combo.bind("<<ComboboxSelected>>", on_selection)

        def on_colonne_selection(event=None):

            for combo, _ in comboboxes:
                combo.destroy()
            comboboxes.clear()

            selected_col = self.colonne_combo.get()
            if selected_col in self.structure_entete and self.structure_entete[selected_col] != {}:
                add_combobox_grid(0, self.structure_entete[selected_col])
            self.colonne_actuelle = get_path()
            self.chemin = get_path()
        self.colonne_combo.bind("<<ComboboxSelected>>", on_colonne_selection)

        def get_path():
            col1 = self.colonne_combo.get()
            selection = [combo.get() for combo, _ in comboboxes if combo.get()]
            chemin = " > ".join([col1] + selection) if col1 else None

            if self.action_selection:
                self.action_selection()
            return chemin

        return get_path




    def get_frame_selection_pack(self, parent_frame):
        """
        Crée une interface en pack pour sélectionner une colonne et ses sous-catégories, et renvoie le chemin sélectionné.

        Args:
        - parent_frame (tk.Frame): Cadre parent où seront placés les widgets.

        Returns:
        - function: Fonction qui retourne le chemin sélectionné.
        """
        self.colonne_combo = ttk.Combobox(parent_frame, values=list(self.structure_entete.keys()), state="readonly")
        self.colonne_combo.pack(padx=5, pady=5, anchor="w")
        self.widgets.append(self.colonne_combo)

        comboboxes = []

        def add_combobox_pack(structure):
            combo = ttk.Combobox(parent_frame, state="readonly")
            combo.pack(padx=5, pady=2, anchor="w")
            combo["values"] = list(structure.keys())
            comboboxes.append((combo, structure))
            self.widgets.append(combo)

            def on_selection(event=None):
                while len(comboboxes) > comboboxes.index((combo, structure)) + 1:
                    widget, _ = comboboxes.pop()
                    widget.destroy()

                selection = combo.get()
                if selection in structure and isinstance(structure[selection], dict) and structure[selection]:
                    add_combobox_pack(structure[selection])

                self.chemin = get_path()

            combo.bind("<<ComboboxSelected>>", on_selection)

        def on_colonne_selection(event=None):
            for combo, _ in comboboxes:
                combo.destroy()
            comboboxes.clear()

            selected_col = self.colonne_combo.get()
            if selected_col in self.structure_entete and self.structure_entete[selected_col]:
                add_combobox_pack(self.structure_entete[selected_col])
            self.colonne_actuelle = get_path()

        self.colonne_combo.bind("<<ComboboxSelected>>", on_colonne_selection)

        def get_path():
            col1 = self.colonne_combo.get()
            selection = [combo.get() for combo, _ in comboboxes if combo.get()]
            self.chemin = " > ".join([col1] + selection) if col1 else None
            return self.chemin

        return get_path

    def grid(self):
        """
        Place tous les widgets de la sélection en utilisant la méthode grid."""
        for widget in self.widgets:
            widget.grid()

    def grid_remove(self):
        """
        Retire tous les widgets de la sélection de la grille."""
        for widget in self.widgets:
            widget.grid_remove()

    def pack(self):
        """
        Place tous les widgets de la sélection en utilisant la méthode pack."""
        for widget in self.widgets:
            widget.pack()

    def pack_forget(self):
        """ Retire tous les widgets de la sélection du pack."""
        for widget in self.widgets:
            widget.pack_forget()
