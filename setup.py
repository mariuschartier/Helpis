from cx_Freeze import setup, Executable

setup(
    name="MonApp",
    version="1.0",
    description="Mon application",
    executables=[Executable("main.py", base=None)]
)
