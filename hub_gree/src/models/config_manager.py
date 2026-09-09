import os
import json
import logging

logger = logging.getLogger(__name__)

class ConfigManager:
    """ Gerencia a leitura e validação das regras externas de negócio. """
    
    # Caminho base adaptável (funciona solto ou compilado via PyInstaller)
    CONFIG_PATH = os.path.join(os.getcwd(), "config", "regras_logistica.json")

    @classmethod
    def carregar_regras(cls) -> dict:
        """ Lê o arquivo JSON. Se não existir, invoca a criação do arquivo padrão. """
        if not os.path.exists(cls.CONFIG_PATH):
            logger.warning("Arquivo de regras não encontrado. Criando base padrão...")
            cls._criar_json_padrao()

        try:
            with open(cls.CONFIG_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            logger.error(f"Erro ao ler JSON de regras: {e}. Usando fallback em memória.")
            return cls._get_regras_padrao()

    @classmethod
    def _criar_json_padrao(cls):
        """ Cria o diretório config/ e o JSON com os dados predefinidos se necessário. """
        os.makedirs(os.path.dirname(cls.CONFIG_PATH), exist_ok=True)
        with open(cls.CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cls._get_regras_padrao(), f, indent=4, ensure_ascii=False)

    @staticmethod
    def _get_regras_padrao() -> dict:
        """ Fallback de segurança com os contratos originais. """
        return {
            "transportadoras": {
                "rodoviario": ["TRANSBUIATTE", "SPEED", "SR LOG", "BUSSOLA", "TODOBRASIL", "AMAZON", "GAB", "BERTOLINI"],
                "cabotagem": ["COSTA BRASIL", "CENTER CARGO", "TECMAR"]
            },
            "regras_faturamento": {
                "TRANSBUIATTE": {
                    "taxa_ponto_extra": 1100,
                    "cidades_especiais": {
                        "MANAUS": {"tipo": "fixo", "valor": 4500}
                    },
                    "padrao": {"tipo": "percentual_valor", "valor": 0.04}
                }
            }
        }