import logging
import os
import sys
from logging.handlers import RotatingFileHandler

def configurar_logger():
    # 1. Define onde o arquivo de log será salvo (Pasta do Usuário do Windows)
    # Isso evita erros de permissão quando o app virar .exe
    pasta_usuario = os.path.expanduser("~")
    pasta_logs = os.path.join(pasta_usuario, "GreeLogisticaLogs")
    
    if not os.path.exists(pasta_logs):
        os.makedirs(pasta_logs)

    arquivo_log = os.path.join(pasta_logs, "gree_hub_erros.log")

    # 2. Configura o Logger Central
    logger = logging.getLogger("GreeHub")
    logger.setLevel(logging.DEBUG) # Captura tudo: INFO, WARNING, ERROR, CRITICAL

    # Evita duplicar logs se a função for chamada mais de uma vez
    if not logger.handlers:
        # Formato: DATA/HORA - NIVEL - [ARQUIVO:LINHA] - MENSAGEM
        formato = logging.Formatter(
            '%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s', 
            datefmt='%d/%m/%Y %H:%M:%S'
        )

        # 3. RotatingFileHandler: Limita o arquivo a 5MB e guarda os 3 mais recentes
        file_handler = RotatingFileHandler(arquivo_log, maxBytes=5*1024*1024, backupCount=3, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(formato)

        # 4. StreamHandler: Continua imprimindo no terminal (VSCode) para você desenvolver
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(logging.INFO)
        console_handler.setFormatter(formato)

        logger.addHandler(file_handler)
        logger.addHandler(console_handler)

    return logger, arquivo_log