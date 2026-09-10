import os
from tkinter import filedialog, messagebox
from views.crm_view import CRMFormatterView
from models.crm_model import CRMModel

class CRMController:
    def __init__(self, parent_frame):
        self.view = CRMFormatterView(parent_frame)
        self.model = CRMModel()
        self.arquivos_selecionados = []
        
        self.view.btnArquivo.configure(command=self.abrir_arquivo)
        self.view.btnExecutar.configure(command=self.executar_processamento)

    def get_view(self):
        return self.view

    def abrir_arquivo(self):
        caminhos = filedialog.askopenfilenames(title="Selecione as planilhas", filetypes=[("Arquivos Excel", "*.xlsx")])
        if caminhos:
            self.arquivos_selecionados = caminhos
            qtd = len(caminhos)
            self.view.lblArquivoStatus.configure(text=f"{qtd} arquivo(s) selecionado(s)", text_color="green")
            self.view.status.configure(text="Status: Analisando arquivos, datas e observações...")
            self.view.update_idletasks()
            
            try:
                datas, mapeamento = self.model.analisar_arquivos(self.arquivos_selecionados)
                self.view.mapeamento_obs = mapeamento
                
                if datas:
                    self.view.combo_datas.configure(values=datas, state="readonly")
                    self.view.combo_datas.set(datas[0])
                    self.view.btnExecutar.configure(state="normal", fg_color=self.view.cor_verde, hover_color="#22953D")
                    self.view.status.configure(text="Status: Arquivos prontos para processamento.")
                    self.view.on_date_change(datas[0])
                else:
                    messagebox.showwarning("Aviso", "Nenhuma data encontrada na Coluna C dos arquivos.")
                    self.view.status.configure(text="Status: Aguardando datas válidas.")
            except Exception as e:
                messagebox.showerror("Erro de Leitura", f"Falha ao processar arquivos: {e}")
                self.view.status.configure(text="Status: Erro na leitura dos arquivos.")

    def executar_processamento(self):
        data_escolhida = self.view.combo_datas.get()
        if not data_escolhida: return

        obs_cli, obs_adi = self.view.get_filtros_selecionados()
        if not obs_cli and not obs_adi and (self.view.checkboxes_obs_cliente or self.view.checkboxes_obs_adic):
            resp = messagebox.askyesno("Aviso", "Você não marcou nenhuma observação como filtro.\nDeseja processar TODAS as linhas da planilha ignorando essa separação?")
            if not resp: return

        self.view.btnExecutar.configure(state="disabled", fg_color="#BDC3C7")
        self.view.btnArquivo.configure(state="disabled")

        total = len(self.arquivos_selecionados)
        sucessos = 0
        arquivo_atual = ""

        try:
            for idx, caminho in enumerate(self.arquivos_selecionados, 1):
                arquivo_atual = os.path.basename(caminho)
                self.view.status.configure(text=f"Status: Processando {idx}/{total} ({arquivo_atual})")

                tipo = self.model.obter_tipo_planilha(caminho)
                sugestao_nome = arquivo_atual.replace(".xlsx", f"_{tipo}_FORMATADO.xlsx")
                caminho_final = filedialog.asksaveasfilename(
                    title=f"Salvar '{arquivo_atual}' como...",
                    initialfile=sugestao_nome,
                    defaultextension=".xlsx",
                    filetypes=[("Arquivos Excel", "*.xlsx")]
                )
                
                if not caminho_final: continue

                self.view.status.configure(text=f"Status: Aplicando regras {tipo} em {arquivo_atual}")
                self.view.update_idletasks()

                progresso_base = (idx - 1) / total
                progresso_por_arq = 1 / total

                def callback_progresso(carga_interna):
                    prog_total = progresso_base + (progresso_por_arq * carga_interna)
                    self.view.progress.set(prog_total)
                    self.view.lbl_porcentagem.configure(text=f"{int(prog_total * 100)}%")
                    self.view.update_idletasks()

                if tipo == "PADRÃO":
                    self.model.formatar_padrao(caminho, data_escolhida, caminho_final, obs_cli, obs_adi, callback_progresso)
                else:
                    # self.model.formatar_gmax(caminho, data_escolhida, caminho_final, obs_cli, obs_adi, callback_progresso)
                    pass
                
                sucessos += 1

            self.view.progress.set(1.0)
            self.view.lbl_porcentagem.configure(text="100%")
            self.view.status.configure(text="Status: Processamento concluído com sucesso!")
            messagebox.showinfo("Sucesso", f"Processamento concluído!\n{sucessos} arquivo(s) salvo(s).")

        except Exception as e:
            messagebox.showerror("Erro Crítico", f"Erro no arquivo: {arquivo_atual}\nErro: {str(e)}")
            self.view.status.configure(text="Status: Processamento interrompido com erro.")
        finally:
            self.view.btnExecutar.configure(state="normal", fg_color=self.view.cor_verde, hover_color="#22953D")
            self.view.btnArquivo.configure(state="normal")