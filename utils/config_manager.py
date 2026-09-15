import json
import os
import logging

logger = logging.getLogger("GreeHub")

class ConfigManager:
    def __init__(self):
        # Aponta diretamente para uma pasta fixa e compartilhada no servidor da empresa (Disco X:)
        self.pasta_app = r"X:\14- Pré-Faturamento\Automações\Projeto piloto - Gree automatização\Laboratório de testes - Hub Gree\hub_gree"
        
        if not os.path.exists(self.pasta_app):
            try:
                os.makedirs(self.pasta_app)
            except Exception as e:
                logger.error(f"Erro ao criar pasta no servidor. Usando pasta local de fallback: {e}")
                self.pasta_app = os.path.expanduser("~")
            
        self.arquivo_config = os.path.join(self.pasta_app, "config.json")
        self.configuracoes = self._carregar_configuracoes()

    def _get_configuracoes_padrao(self):
        return {
            "frete_transbuiatte_base_manaus": 4500.0,
            "frete_transbuiatte_acrescimo": 1100.0,
            "solicitacao_padrao_consolidador": "2026080168",
            "dpi_padrao_conversao": "150 (Médio)",
            "tema_aparencia": "light",
            "tema_cor_ctk": "blue"
        }

    def _carregar_configuracoes(self):
        if not os.path.exists(self.arquivo_config):
            padrao = self._get_configuracoes_padrao()
            self._salvar(padrao)
            return padrao
        
        try:
            with open(self.arquivo_config, 'r', encoding='utf-8') as f:
                config_atual = json.load(f)
                padrao = self._get_configuracoes_padrao()
                teve_atualizacao = False
                for chave, valor in padrao.items():
                    if chave not in config_atual:
                        config_atual[chave] = valor
                        teve_atualizacao = True
                
                if teve_atualizacao:
                    self._salvar(config_atual)
                    
                return config_atual
        except Exception as e:
            logger.error(f"Erro ao ler config.json no servidor: {e}")
            return self._get_configuracoes_padrao()

    def _salvar(self, config_dict):
        try:
            with open(self.arquivo_config, 'w', encoding='utf-8') as f:
                json.dump(config_dict, f, indent=4, ensure_ascii=False)
        except Exception as e:
            logger.error(f"Erro ao salvar config.json no servidor: {e}")

    def get(self, chave):
        return self.configuracoes.get(chave, self._get_configuracoes_padrao().get(chave))

    def set(self, chave, valor):
        self.configuracoes[chave] = valor
        self._salvar(self.configuracoes)