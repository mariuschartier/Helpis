from cx_Freeze import setup, Executable
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
    executables=[Executable("main.py")],
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