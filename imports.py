"""
### This file is used to import all the necessary libraries for the project."""
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
        if alias == None:
            alias = package
        # Tente d'importer le package
        mod = __import__(package)
        print(f"{package} est déjà installé.")
    except ImportError:
        print(f"{package} n'est pas installé. Installation en cours...")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
        # Réessaie d'importer le package après l'installation
    # try:
    #         mod = __import__(alias)
    # except Exception as e:
    #      print(f"Erreur lors de l'importation de {package} : {e}")


    # if alias:
    #     # Utilise globals() pour définir l'alias
    #     globals()[alias] = mod
    #     # print(f"{package} importé sous l'alias '{alias}'.")
    #     return mod
        

install_and_import('pandas')
install_and_import('json')
install_and_import('tkinter')
install_and_import('threading')
install_and_import('os')
install_and_import('pathlib')
install_and_import('openpyxl')
install_and_import('typing')
install_and_import('datetime')
install_and_import('bs4')
install_and_import('chardet')
install_and_import('scipy')
install_and_import('xlrd')
install_and_import('pywin32','win32com')
install_and_import('matplotlib')
install_and_import('jsonpickle')
install_and_import('xlsxwriter')             
install_and_import('ttkbootstrap')             


