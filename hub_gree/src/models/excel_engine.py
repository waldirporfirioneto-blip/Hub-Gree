import os
import openpyxl
import logging

logger = logging.getLogger(__name__)

class ExcelEngine:
    """ Manipula I/O de planilhas. Independente de UI. """
    
    @staticmethod
    def ler_cabecalhos(caminho_arquivo: str) -> list:
        """ Lê os cabeçalhos da primeira linha atenuando consumo de memória. """
        try:
            wb = openpyxl.load_workbook(caminho_arquivo, read_only=True, data_only=True)
            ws = wb.active
            # Utiliza generator para eficiência (lazy evaluation)
            primeira_linha = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
            wb.close()
            return [str(celula).strip().lower() for celula in primeira_linha if celula]
        except Exception as e:
            logger.error(f"Falha ao ler cabeçalho de {caminho_arquivo}: {e}")
            raise

    @staticmethod
    def aplicar_formatacao_padrao(caminho_origem: str, caminho_destino: str):
        
        # operando unicamente no openpyxl, sem pop-ups do tkinter.
        pass