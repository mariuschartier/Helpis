
from structure.Fichier import Fichier
from structure.Feuille import Feuille
from datetime import datetime
from openpyxl import Workbook
import pandas as pd

import subprocess
import sys
# import pkg_resources

def install_and_import(package, alias=None):
    """Installe un package Python et l'importe sous un alias spécifié. 
    Si le package est déjà installé, il l'importe directement.

    Args:
        package (str): Le nom du package à installer et importer.
        alias (str, optional): L'alias sous lequel importer le package.
    """
    try:
        # Tente d'importer le package
        mod = __import__(package)
        print(f"{package} est déjà installé.")
    except ImportError:
        print(f"{package} n'est pas installé. Installation en cours...")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
        # Réessaie d'importer le package après l'installation
        mod = __import__(package)

    if alias:
        # Utilise globals() pour définir l'alias
        globals()[alias] = mod
        print(f"{package} importé sous l'alias '{alias}'.")
        return mod
        

