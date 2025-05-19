import imports
from structure.Fichier import Fichier
from structure.Feuille import Feuille
from tests import fonctions as fn
from dataclasses import dataclass,asdict
from front import multi_page as mp

# main.py
import sys

def main():
    """Point d'entrée principal de l'application."""
    print("Bienvenue dans mon projet Python 🚀")
    
    # Exemple : appeler d'autres modules ici
    # from core.moteur import demarrer
    # demarrer()

if __name__ == "__main__":
    try:
        main()
        
        app = mp.MultiPageApp()
        app.mainloop()
        # f1 = Fichier("results/Mesures.xlsx")
        # f1_1 = Feuille(f1,"opti",2)
        # fn.moyenne_par_semaine(f1_1)
    except Exception as e:
        print(f"Erreur fatale : {e}", file=sys.stderr)                                                  
        sys.exit(1)

        

        