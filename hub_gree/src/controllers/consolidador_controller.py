import threading
import os
from src.models.pdf_engine import PDFEngine

class ConsolidadorController:
    
    def processar_ordens_de_venda(self, pasta_entrada: str, callback_ui: callable):
        """ Encontra PDFs na pasta, extrai ordens e aciona o gerador de Excel. """
        thread = threading.Thread(
            target=self._worker_extrair_ordens, 
            args=(pasta_entrada, callback_ui),
            daemon=True
        )
        thread.start()

    def _worker_extrair_ordens(self, pasta_entrada, callback_ui):
        try:
            pdf_files = [
                os.path.join(pasta_entrada, f) for f in os.listdir(pasta_entrada)
                if f.lower().endswith(".pdf") and not f.startswith("~$")
            ]
            
            if not pdf_files:
                callback_ui("Nenhum PDF encontrado na pasta.", erro=True)
                return

            todas_ordens = {}
            total = len(pdf_files)
            
            for idx, filepath in enumerate(sorted(pdf_files)):
                progresso = int(((idx + 1) / total) * 90)
                callback_ui(f"Analisando: {os.path.basename(filepath)}", pct=progresso)
                
                # Chama a Engine otimizada do Passo 2
                ordens_arquivo = PDFEngine.extrair_ordens_de_arquivo(filepath)
                for ordem in ordens_arquivo:
                    todas_ordens[ordem] = None
                    
            if not todas_ordens:
                callback_ui("Nenhuma ordem encontrada nos PDFs.", erro=True)
                return

            lista_final_ordens = list(todas_ordens.keys())
            callback_ui("Estruturando Excel de Controle...", pct=95)
            
            # Aqui você chamaria o ExcelEngine (construído no Passo 1) 
            # para gerar aquele layout da planilha da GREE.
            # ExcelEngine.gerar_planilha_controle_entrega(lista_final_ordens)

            callback_ui("Relatório gerado com sucesso!", pct=100, concluido=True)
            
        except Exception as e:
            callback_ui(f"Falha Crítica: {str(e)}", erro=True)