import sys
from cx_Freeze import setup, Executable

build_exe_options = {"packages": ["queue","requests","openpyxl","time","webbrowser","threading","logging","PIL","CTkMessagebox"], "includes": ["tkinter","customtkinter"]}

base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="Rub Import",
    version="1.0",
    description="Aplicação integrada com o CRM Rubeus, com o objetivo de automatizar importações de Curso e Oferta de curso em massa.",
    options={"build_exe": build_exe_options},
    executables=[Executable("App.py", base=base, icon='./assets/icon4.ico', target_name='Rub Import')]
)