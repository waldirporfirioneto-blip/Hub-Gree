from tkinter import filedialog, messagebox
from views.cotacao_view import CotacaoView
from models.cotacao_model import CotacaoModel

class CotacaoController:
    def __init__(self, parent_frame):
        self.model = CotacaoModel()
        self.view = CotacaoView(parent_frame)
        self.view.btn_carregar.configure(command=self.carregar_arquivo)
        self.view.btn_gerar.configure(command=self.gerar_planilha)

    def get_view(self):
        return self.view

    def carregar_arquivo(self):
        caminho_entrada = filedialog.askopenfilename(
            title="Selecione a base de dados (Excel)",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
        )
        if not caminho_entrada: return

        try:
            itens, visao, qtd_cargas = self.model.processar_base_dados(caminho_entrada)
            self.view.atualizar_status(f"Sucesso! Visão Identificada: {visao} | Cargas CIF: {qtd_cargas}", "#27AE60")
            self.view.preencher_lista(itens)
            self.view.ativar_botao_gerar()

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler o arquivo:\n\n{str(e)}")
            self.view.desativar_botao_gerar()
            self.view.atualizar_status("Erro no processamento do arquivo", "red")

    def gerar_planilha(self):
        arquivo_saida = filedialog.asksaveasfilename(
            title="Onde deseja salvar a Cotação Formatada?",
            defaultextension=".xlsx",
            filetypes=[("Arquivos Excel", "*.xlsx")],
            initialfile="COTACAO_FORMATADA.xlsx"
        )
        if not arquivo_saida: return

        try:
            ordem_desejada = self.view.obter_ordem_lista()
            
            self.model.gerar_planilha(arquivo_saida, ordem_desejada)
            
            messagebox.showinfo("Sucesso!", "Planilha Gerada com Sucesso!\n\nAs abas Rodoviário, Cabotagem e Cargas CIF foram salvas perfeitamente.")

        except Exception as e:
            messagebox.showerror("Erro Crítico", f"Ocorreu um erro ao salvar o arquivo Excel:\n\n{str(e)}")