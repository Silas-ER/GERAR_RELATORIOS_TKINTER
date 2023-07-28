import sys
from cx_Freeze import setup, Executable

# Inclua aqui as bibliotecas que você está usando
build_exe_options = {"packages": ["pandas", "pyodbc", "pyautogui", "time", "tkinter", "tkcalendar", "datetime", "PIL"], 
                     "include_files": ["img\\logo.png"]}

base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="Gerador de Relatorios",
    version="0.1",
    description="Cria Relatórios de custo de barcos através da conexão com o banco de dados sql fazendo filtros com scripts e importando para o excel",
    options={"build_exe": build_exe_options},
    executables=[Executable("custo_barcos.py", base=base)]
)

