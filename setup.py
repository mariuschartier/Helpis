from cx_Freeze import setup, Executable


import os
import sys
from pathlib import Path

def prepare_dossiers():
    # Vérifier si le script est empaqueté
    if hasattr(sys, '_MEIPASS'):
        base_dir = Path(sys._MEIPASS)
    else:
        base_dir = Path(__file__).parent

    sauvegardes_dir = base_dir / 'sauvegardes'

    # Créer les dossiers nécessaires
    (sauvegardes_dir / 'sauvegardes_tests').mkdir(parents=True, exist_ok=True)
    (sauvegardes_dir / 'results').mkdir(parents=True, exist_ok=True)
    (sauvegardes_dir / 'data').mkdir(parents=True, exist_ok=True)

    print(f"Dossiers créés dans : {sauvegardes_dir}")




# Si votre application est une GUI, utilisez "Win32GUI"
base = "Win32GUI"

buildOptions = {
    "packages": ["pandas",
                "json",
                "tkinter",
                "threading",
                "os",
                "pathlib",
                "openpyxl",
                "typing",
                "datetime",
                "bs4",
                "chardet",
                "scipy",
                "xlrd",
                "win32com",
                "matplotlib",
                "jsonpickle",
                "xlsxwriter",
                "ttkbootstrap",
                "itertools",],  # Ajoutez ici les modules nécessaires
    "excludes": [],  # Modules à exclure si besoin
    "include_files": [("logo.ico", "logo.ico")],  # Ressources à inclure
}

setup(
    name="MonApplication",
    version="1.0",
    description="Description de mon app",
    options={"build_exe": buildOptions},
    executables=[Executable("main.py", base=base, icon="logo.ico")],
)




# setup(
#     name="MonApplication",
#     version="1.0",
#     description="Description de mon app",
#     executables=[Executable("imports.py"),Executable("tmp.py")],
# )


# setup(
#     name="MonApplication",
#     version="1.0",
#     description="Description de mon app",
#     executables=[Executable("imports.py")],
# )