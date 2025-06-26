# import imports as imp
from structure.Fichier import Fichier
from structure.Feuille import Feuille
import fonctions as fn
from dataclasses import dataclass,asdict
from front import multi_page as mp

# main.py
import sys

def main():
    

    """Point d'entrée principal de l'application."""
    print("Bienvenue dans mon projet Python 🚀")
    


if __name__ == "__main__":
    try:
        # imp.import_bibli()          
        main()
        
        app = mp.MultiPageApp()
        app.mainloop()
    except Exception as e:
        print(f"Erreur fatale : {e}", file=sys.stderr)                                                  
        sys.exit(1)

        

        