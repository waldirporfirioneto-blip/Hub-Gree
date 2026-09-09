import threading
from src.models.excel_engine import ExcelEngine

class CRMController:
    def __init__(self):
        self.engine = ExcelEngine()

    def processar_arquivos_lote(self, caminhos: tuple, ui_callback: callable):
        """ 
        Inicia thread dedicada para I/O intensivo. 
        O callback atualiza a View de forma segura.
        """
        thread = threading.Thread(
            target=self._worker_processamento, 
            args=(caminhos, ui_callback),
            daemon=True
        )
        thread.start()

    def _worker_processamento(self, caminhos, ui_callback):
        try:
            for idx, caminho in enumerate(caminhos, 1):
                ui_callback(f"Processando {idx}/{len(caminhos)}...")
                
                cabecalhos = self.engine.ler_cabecalhos(caminho)
                
                # Validação de regras de negócio (G-MAX vs Padrão)
                tipo_planilha = "G-MAX" if any(h in cabecalhos for h in ["código unidade", "unidade"]) else "PADRÃO"
                
                # Delega ao Model a execução das transformações
                # self.engine.aplicar_formatacao_padrao(caminho, caminho_saida)
                
            ui_callback("Processamento concluído com sucesso!", concluido=True)
        except Exception as e:
            ui_callback(f"Erro Crítico: {str(e)}")