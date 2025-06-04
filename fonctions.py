
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