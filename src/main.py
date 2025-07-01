import ttkbootstrap as ttk
from tkinter import TclError
from src.gui import AppGUI
import os

def main():
    """Função principal que inicia o aplicativo."""
    root = ttk.Window()
    root.minsize(900, 720)

    # Define o caminho para o arquivo de ícone dentro da pasta de assets
    icon_path = os.path.join('assets', 'rbc_logo.ico')

    try:
        root.iconbitmap(icon_path)
    except TclError: # Use TclError que foi importado
        print(f"Aviso: Ícone '{icon_path}' não encontrado. Usando ícone padrão.")
    # --- FIM DA MODIFICAÇÃO CORRIGIDA ---

    root.state("zoomed")
    root.bind("<Escape>", lambda e: root.attributes("-fullscreen", False))
    root.title("Projeto RBCTur")
    root.geometry("1024x768")

    AppGUI(root)

    root.mainloop()

if __name__ == "__main__":
    main()