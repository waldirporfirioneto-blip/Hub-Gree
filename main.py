import sys
import subprocess
import traceback
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk

# ==========================================
# VERIFICAÇÃO E INSTALAÇÃO AUTOMÁTICA
# ==========================================
# Executa a checagem APENAS se estiver rodando o código-fonte solto (VSCode/Terminal)
if not getattr(sys, 'frozen', False):
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

# Importações dos controladores e logger após a trava
from controllers.main_controller import MainController
from utils.logger import configurar_logger

# ==========================================
# TRATAMENTO GLOBAL DE EXCEÇÕES
# ==========================================
logger, arquivo_log = configurar_logger()

def manipulador_excecoes_global(tipo_erro, valor_erro, traceback_erro):
    mensagem_log = "".join(traceback.format_exception(tipo_erro, valor_erro, traceback_erro))
    logger.critical(f"Erro Fatal Não Tratado:\n{mensagem_log}")
    
    root = tk.Tk()
    root.withdraw()
    mensagem_alerta = (
        "Ocorreu um erro inesperado no sistema.\n\n"
        f"Detalhes técnicos foram salvos para a equipe de TI no arquivo:\n{arquivo_log}\n\n"
        f"Erro base: {valor_erro}"
    )
    messagebox.showerror("Erro Crítico - Gree Logística Hub", mensagem_alerta)
    root.destroy()

sys.excepthook = manipulador_excecoes_global

# ==========================================
# INICIALIZAÇÃO DO SISTEMA
# ==========================================
def main():
    logger.info("Iniciando o Gree Logística Hub...")
    
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")
    
    try:
        import ctypes
        myappid = 'gree.logistica.hub.v1'
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except Exception as e:
        logger.warning(f"Falha ao definir App ID do Windows: {e}")

    app = MainController()
    app.iniciar()
    logger.info("Gree Logística Hub encerrado pelo usuário.")

if __name__ == "__main__":
    main()