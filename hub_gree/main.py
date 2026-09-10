import sys
import subprocess
import customtkinter as ctk

def checar_dependencias():
    dependencias = {
        "pandas": "pandas",
        "openpyxl": "openpyxl",
        "customtkinter": "customtkinter",
        "PIL": "Pillow",
        "win32com": "pywin32",
        "pypdf": "pypdf",
        "docx": "python-docx",
        "pypdfium2": "pypdfium2"
    }
    for modulo, pacote in dependencias.items():
        try:
            __import__(modulo)
        except ImportError:
            print(f"Instalando pacote necessário: {pacote}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", pacote])

checar_dependencias()

from controllers.main_controller import MainController

def main():
   
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")
    
    
    try:
        import ctypes
        myappid = 'gree.logistica.hub.v1'
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except Exception:
        pass
    app = MainController()
    app.iniciar()

if __name__ == "__main__":
    main()