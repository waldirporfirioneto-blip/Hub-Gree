import os
import re
import logging
import pypdf

logger = logging.getLogger(__name__)

class PDFEngine:
    """ Engine otimizada para manipulação e extração de dados de PDFs. """
    
    # Pré-compilamos a expressão regular uma única vez na subida do módulo.
    # Padrões buscados: S + 8 dígitos, 5 + 8 dígitos, 100 + 6 dígitos, TPZ/TPR + 6 dígitos.
    REGEX_ORDENS = re.compile(r'\b([S5]\d{8}|100\d{6}|TPZ\d{6}|TPR\d{6})\b', re.IGNORECASE)

    @classmethod
    def extrair_ordens_de_arquivo(cls, filepath: str) -> list:
        """
        Lê o PDF sob demanda (página a página) para não estourar a memória.
        Retorna uma lista de códigos únicos extraídos.
        """
        # Usamos dict para garantir chaves únicas mantendo a ordem de inserção (O(1) lookup)
        ordens_encontradas = {} 
        
        try:
            with open(filepath, "rb") as f:
                reader = pypdf.PdfReader(f, strict=False)
                for page in reader.pages:
                    text = page.extract_text()
                    if text:
                        # Extrai imediatamente e libera o texto da RAM
                        cls._processar_regex_em_texto(text, ordens_encontradas)
        except Exception as e:
            logger.warning(f"Erro ao ler PDF internamente ({filepath}): {e}. Tentando fallback no nome.")

        # Fallback: Se não encontrou nada no conteúdo, tenta ler no nome do arquivo
        if not ordens_encontradas:
            nome_sem_ext = os.path.splitext(os.path.basename(filepath))[0]
            cls._processar_regex_em_texto(nome_sem_ext, ordens_encontradas)

        return list(ordens_encontradas.keys())

    @classmethod
    def _processar_regex_em_texto(cls, texto: str, dict_armazenamento: dict):
        """ Aplica a regex e padroniza as ordens encontradas. """
        matches = cls.REGEX_ORDENS.findall(texto)
        for match in matches:
            codigo = match.upper()
            
            # Regra de negócio: Padronizar prefixo '5' para 'S'
            if codigo.startswith("5"):
                codigo = "S" + codigo[1:]
                
            # Salva no dicionário. O valor None é irrelevante, usamos apenas a chave.
            dict_armazenamento[codigo] = None

    @staticmethod
    def mesclar_arquivos_pdf(arquivos_entrada: list, caminho_saida: str, callback_progresso=None):
        """ Junta os PDFs sem estourar limite de I/O de arquivos abertos simultaneamente. """
        writer = pypdf.PdfWriter()
        total = len(arquivos_entrada)
        
        for idx, file_path in enumerate(arquivos_entrada):
            if os.path.normpath(file_path) == os.path.normpath(caminho_saida):
                continue
            
            try:
                # O reader carrega os ponteiros, o writer salva a referência
                reader = pypdf.PdfReader(file_path, strict=False)
                for page in reader.pages:
                    writer.add_page(page)
            except Exception as e:
                logger.error(f"Erro ao anexar {file_path}: {e}")
                
            if callback_progresso:
                callback_progresso(idx + 1, total, f"Anexando {os.path.basename(file_path)}")

        # Salva o arquivo final numa única operação de gravação
        with open(caminho_saida, "wb") as f_out:
            writer.write(f_out)
        writer.close()