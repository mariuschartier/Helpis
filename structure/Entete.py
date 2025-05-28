from openpyxl.styles import Font

class Entete:
    """ Représente l'entête d'une feuille Excel, permettant de manipuler les métadonnées de l'entête."""
    def __init__(self, feuille, entete_debut=0, entete_fin=1, nb_colonnes_secondaires=0, ligne_unite=1,structure = {}):
        self.feuille = feuille
        self.entete_debut = entete_debut
        self.entete_fin = entete_fin
        self.nb_colonnes_secondaires = nb_colonnes_secondaires
        self.ligne_unite = ligne_unite
        self.taille_entete = self.entete_fin - self.entete_debut
        self.structure = structure
        self.placement_colonne = self.set_position()
        self.df_entete = feuille.df.iloc[entete_debut:entete_fin + 1]  # Attention : entete_fin inclus

    def copier_dans_ws(self, ws, start_row=1, start_col=1, style_gras=True):
        """Copier l'entête dans une feuille openpyxl, avec options style."""
        for row_idx in range(self.df_entete.shape[0]):
            for col_idx in range(self.df_entete.shape[1]):
                valeur = self.df_entete.iloc[row_idx, col_idx]
                cell = ws.cell(row=start_row + row_idx, column=start_col + col_idx, value=valeur)
                if style_gras:
                    cell.font = Font(bold=True)

    def get_nb_lignes(self):
        """Retourner le nombre total de lignes d'entête."""
        return self.entete_fin - self.entete_debut + 1

    def get_lignes(self):
        """Retourner l'entête sous forme de liste de listes."""
        return self.df_entete.values.tolist()

    def get_unite(self):
        """Retourne la ligne des unités (ligne_unite) si disponible."""
        if self.entete_debut <= self.ligne_unite <= self.entete_fin:
            return self.df_entete.iloc[self.ligne_unite - self.entete_debut].tolist()
        else:
            return None

    def __str__(self):
        return (f"Entête '{self.feuille.nom}' de la ligne {self.entete_debut} à {self.entete_fin}, "
                f"avec {self.nb_colonnes_secondaires} colonnes secondaires, unité à la ligne {self.ligne_unite}.")
    

    def set_position(self):
        """
        Génère un dictionnaire avec les chemins hiérarchiques des colonnes comme clés
        et leurs indices correspondants comme valeurs.
        Exemple : {"Col 1 > Scol 2 > Sscol 1": 1, "Col 1 > Scol 2 > Sscol 2": 2}
        """
        positions = {}

        def parcourir_structure(structure:dict, chemin="", index=0):
            for cle, valeur in structure.items():
                chemin_actuel = f"{chemin} > {cle}" if chemin else cle

                if isinstance(valeur, dict) and valeur:  # Si c'est un dictionnaire non vide
                    index = parcourir_structure(valeur, chemin_actuel, index)
                else:  # Si c'est une feuille (fin de la hiérarchie)
                    positions[chemin_actuel] = index
                    index += 1

            return index
        parcourir_structure(self.structure)

        return positions
        

    def une_ligne(self)->list: 
        """
        revoi une version en une seule ligne de l'entete
        """
        return self.placement_colonne.keys()
            

    def maj_entete(self, entete_debut: int=None,
                        entete_fin: int=None,
                        nb_colonnes_secondaires: int=None,
                        ligne_unite : int=None,
                        structure : dict=None):
        """
        Met à jour l'entête de la feuille cible avec les données de l'entête actuelle.
        """
        if isinstance(entete_debut, int):
            self.entete_debut = entete_debut
        if isinstance(entete_fin, int):
            self.entete_fin = entete_fin
        if isinstance(nb_colonnes_secondaires, int):
            self.nb_colonnes_secondaires = nb_colonnes_secondaires
        if isinstance(ligne_unite, int):
            self.ligne_unite = ligne_unite
        if isinstance(structure, dict):
            self.structure = structure
        self.taille_entete = self.entete_fin - self.entete_debut
        self.placement_colonne = self.set_position()
