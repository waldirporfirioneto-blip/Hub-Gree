import os
import sys


def obter_caminho_recurso(nome_arquivo):
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        caminho_interno = os.path.join(sys._MEIPASS, nome_arquivo)
        if os.path.exists(caminho_interno):
            return caminho_interno

    diretorio_script = os.path.dirname(os.path.abspath(__file__))
    diretorio_raiz = os.path.dirname(diretorio_script)
    caminho_direto = os.path.join(diretorio_raiz, nome_arquivo)
    
    if os.path.exists(caminho_direto):
        return caminho_direto

    return nome_arquivo

# Constantes Globais de Recursos
CAMINHO_LOGO = obter_caminho_recurso("gree.png")
CAMINHO_ICO = obter_caminho_recurso("gree.ico")
CAMINHO_LOGO_COR = obter_caminho_recurso("logo gree colorida.png")