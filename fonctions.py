
from structure.Fichier import Fichier
from structure.Feuille import Feuille
from datetime import datetime
from openpyxl import Workbook
import pandas as pd



def to_int(val):
    try:
            val = int(val)
    except ValueError:
        print("val is not convertible to an integer")
    return val


def is_file_locked(filepath):
    try:
        # Essayer d'ouvrir le fichier en mode écriture
        with open(filepath, 'r+'):
            return False  # Le fichier n'est pas verrouillé
    except IOError:
        return True  # Le fichier est utilisé ou verrouillé