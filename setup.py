import sys
import os
from cx_Freeze import setup, Executable

# --- CAMINHOS E CONFIGURAÇÕES ---
main_script = os.path.join("src", "main.py")

# O ícone está dentro de src/assets.
icon_path = os.path.join("src", "assets", "rbc_logo.ico")
assets_folder = os.path.join("src", "assets")

# --- OPÇÕES DE BUILD  ---
build_exe_options = {
    # Lista de módulos que podem não ser detectados automaticamente
    "includes": [
        "tkinter", "ttkbootstrap", "queue", "threading", "ttkbootstrap.utility"
    ],
    # Lista de pacotes a serem incluídos.
    "packages": [
        "os", "gspread", "pandas", "google", "requests", "ttkbootstrap", "src"
    ],

    "include_files": [
        (assets_folder, "assets")
    ]
}

# Base para aplicações com interface gráfica no Windows
base = "Win32GUI" if sys.platform == "win32" else None

# --- SETUP ---
setup(
    name="SincronizadorRBC",
    version="1.0",
    description="Sincronizador de Contatos com Google Sheets",
    options={"build_exe": build_exe_options},
    executables=[Executable(main_script, base=base, icon=icon_path)]
)