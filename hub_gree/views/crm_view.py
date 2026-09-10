import customtkinter as ctk
import tkinter as tk
from PIL import Image
from utils.config_app import CAMINHO_LOGO

class CRMFormatterView(ctk.CTkFrame):
    def __init__(self, master=None):
        super().__init__(master, fg_color="transparent")
        self.cor_azul = "#0057B8"
        self.cor_verde = "#2EAF4A"
        self.checkboxes_obs_cliente = []
        self.checkboxes_obs_adic = []
        self.selecionados_cliente = set()
        self.selecionados_adic = set()
        self.mapeamento_obs = [] # Atualizado pelo controller
        
        self.criar_interface()

    def criar_interface(self):
        header = ctk.CTkFrame(self, fg_color=self.cor_azul, corner_radius=0)
        header.pack(fill="x")

        try:
            logo = ctk.CTkImage(light_image=Image.open(CAMINHO_LOGO), size=(130, 30))
            lblLogo = ctk.CTkLabel(header, image=logo, text="")
            lblLogo.pack(side="left", padx=25, pady=12)
        except: pass

        titulo = ctk.CTkLabel(header, text="CRM Formatação de Planilhas", text_color="white", font=("Segoe UI", 18, "bold"))
        titulo.pack(side="left", padx=15, pady=12)

        body = ctk.CTkFrame(self, fg_color="white", corner_radius=15)
        body.pack(padx=25, pady=20, fill="both", expand=True)

        container_colunas = ctk.CTkFrame(body, fg_color="transparent")
        container_colunas.pack(fill="both", expand=True, padx=20, pady=(15, 0))

        col_esq = ctk.CTkFrame(container_colunas, fg_color="transparent")
        col_esq.pack(side="left", fill="both", expand=True, padx=(0, 15))

        lbl1 = ctk.CTkLabel(col_esq, text=" 📂 Selecione os arquivos Excel (.xlsx)", font=("Segoe UI", 16, "bold"), text_color="#1C2D42")
        lbl1.pack(anchor="w", pady=(10, 5))

        self.btnArquivo = ctk.CTkButton(col_esq, text="Selecionar Arquivos (Lote)", height=45, corner_radius=8, font=("Segoe UI", 13, "bold"))
        self.btnArquivo.pack(fill="x")

        self.lblArquivoStatus = ctk.CTkLabel(col_esq, text="Nenhum arquivo selecionado", text_color="gray")
        self.lblArquivoStatus.pack(anchor="w", pady=10)

        lbl2 = ctk.CTkLabel(col_esq, text=" 📅 Escolha a data de Programação", font=("Segoe UI", 16, "bold"), text_color="#1C2D42")
        lbl2.pack(anchor="w", pady=(80, 5))

        self.combo_datas = ctk.CTkComboBox(col_esq, values=[], height=40, state="disabled", command=self.on_date_change)
        self.combo_datas.pack(fill="x")

        lbl_dica = ctk.CTkLabel(col_esq, text="Dica: O sistema lerá todas as datas e observações disponíveis nos arquivos\nselecionados automaticamente.", font=("Segoe UI", 12, "italic"), text_color="gray60", justify="left")
        lbl_dica.pack(anchor="w", pady=(30, 0))

        divisoria = ctk.CTkFrame(container_colunas, width=2, fg_color="#E0E6ED")
        divisoria.pack(side="left", fill="y", padx=18, pady=10)

        col_dir = ctk.CTkFrame(container_colunas, fg_color="transparent")
        col_dir.pack(side="right", fill="both", expand=True, padx=(15, 0))

        lbl3 = ctk.CTkLabel(col_dir, text=" ✔ Filtros Dinâmicos de Observação", font=("Segoe UI", 16, "bold"), text_color="#1C2D42")
        lbl3.pack(anchor="w", pady=(10, 5))

        frames_obs_container = ctk.CTkFrame(col_dir, fg_color="transparent")
        frames_obs_container.pack(fill="both", expand=True)

        frame_cliente = ctk.CTkFrame(frames_obs_container, fg_color="transparent")
        frame_cliente.pack(side="left", fill="both", expand=True, padx=(0, 5))
        ctk.CTkLabel(frame_cliente, text="Obs. Cliente", font=("Segoe UI", 13, "bold"), text_color="#333333").pack(anchor="w")
        self.chk_all_cliente = ctk.CTkCheckBox(frame_cliente, text="Selecionar Todas", font=("Segoe UI", 12, "bold"), command=self.toggle_all_cliente)
        self.chk_all_cliente.pack(anchor="w", pady=(5, 5))
        self.scroll_obs_cliente = ctk.CTkScrollableFrame(frame_cliente, fg_color="#F0F2F5", corner_radius=8)
        self.scroll_obs_cliente.pack(fill="both", expand=True, pady=(0, 10))

        frame_adic = ctk.CTkFrame(frames_obs_container, fg_color="transparent")
        frame_adic.pack(side="left", fill="both", expand=True, padx=(5, 0))
        ctk.CTkLabel(frame_adic, text="Obs. Adicionais", font=("Segoe UI", 13, "bold"), text_color="#333333").pack(anchor="w")
        self.chk_all_adic = ctk.CTkCheckBox(frame_adic, text="Selecionar Todas", font=("Segoe UI", 12, "bold"), command=self.toggle_all_adic)
        self.chk_all_adic.pack(anchor="w", pady=(5, 5))
        self.scroll_obs_adic = ctk.CTkScrollableFrame(frame_adic, fg_color="#F0F2F5", corner_radius=8)
        self.scroll_obs_adic.pack(fill="both", expand=True, pady=(0, 10))

        rodape_execucao = ctk.CTkFrame(body, fg_color="transparent")
        rodape_execucao.pack(fill="x", side="bottom", padx=25, pady=(25, 20), before=container_colunas)

        self.btnExecutar = ctk.CTkButton(
            rodape_execucao, width=260, height=55, corner_radius=10,
            fg_color="#BDC3C7", text_color="white", text="EXECUTAR FORMATAÇÃO EM LOTE",
            font=("Segoe UI", 16, "bold"), state="disabled"
        )
        self.btnExecutar.pack(pady=(15, 10), fill="x")

        progress_frame = ctk.CTkFrame(rodape_execucao, fg_color="transparent")
        progress_frame.pack(fill="x", pady=(0, 5))
        self.progress = ctk.CTkProgressBar(progress_frame, height=14, progress_color=self.cor_verde, fg_color="#E0E6ED")
        self.progress.pack(side="left", fill="x", expand=True, padx=(0, 15))
        self.progress.set(0)
        self.lbl_porcentagem = ctk.CTkLabel(progress_frame, text="0%", font=("Segoe UI", 14, "bold"), text_color=self.cor_azul)
        self.lbl_porcentagem.pack(side="right")

        info_frame = ctk.CTkFrame(rodape_execucao, fg_color="transparent")
        info_frame.pack(fill="x")
        self.status = ctk.CTkLabel(info_frame, text="Status: Aguardando arquivos...", text_color="gray40")
        self.status.pack(side="left")
        ctk.CTkLabel(info_frame, text="Desenvolvido para Operação GREE LOGÍSTICA", text_color="gray50", font=("Segoe UI", 10, "italic")).pack(side="right")

    # Métodos puros da View original preservados 
    def on_date_change(self, nova_data):
        self.selecionados_cliente.clear()
        self.selecionados_adic.clear()
        for row in self.mapeamento_obs:
            if row['data'] == nova_data:
                if "nf" in row['cliente'].lower(): self.selecionados_cliente.add(row['cliente'])
                if "nf" in row['adic'].lower(): self.selecionados_adic.add(row['adic'])
        self.chk_all_cliente.deselect()
        self.chk_all_adic.deselect()
        self.atualizar_painel_filtros()

    def cb_cliente_clicked(self):
        self.selecionados_cliente = {cb.cget("text") for cb in self.checkboxes_obs_cliente if cb.get()}
        self.atualizar_painel_filtros()

    def cb_adic_clicked(self):
        self.selecionados_adic = {cb.cget("text") for cb in self.checkboxes_obs_adic if cb.get()}
        self.atualizar_painel_filtros()

    def toggle_all_cliente(self):
        estado = self.chk_all_cliente.get()
        for cb in self.checkboxes_obs_cliente:
            if estado: cb.select()
            else: cb.deselect()
        self.cb_cliente_clicked()

    def toggle_all_adic(self):
        estado = self.chk_all_adic.get()
        for cb in self.checkboxes_obs_adic:
            if estado: cb.select()
            else: cb.deselect()
        self.cb_adic_clicked()

    def atualizar_painel_filtros(self):
        data_escolhida = self.combo_datas.get()
        if not data_escolhida: return
        valid_clientes, valid_adics = set(), set()
        for row in self.mapeamento_obs:
            if row['data'] == data_escolhida:
                if not self.selecionados_adic or row['adic'] in self.selecionados_adic:
                    valid_clientes.add(row['cliente'])
                if not self.selecionados_cliente or row['cliente'] in self.selecionados_cliente:
                    valid_adics.add(row['adic'])

        for widget in self.scroll_obs_cliente.winfo_children(): widget.destroy()
        self.checkboxes_obs_cliente.clear()
        if valid_clientes:
            for obs in sorted(list(valid_clientes)):
                cb = ctk.CTkCheckBox(self.scroll_obs_cliente, text=obs, font=("Segoe UI", 12), command=self.cb_cliente_clicked)
                cb.pack(anchor="w", padx=10, pady=5)
                if obs in self.selecionados_cliente: cb.select()
                self.checkboxes_obs_cliente.append(cb)
        else:
            ctk.CTkLabel(self.scroll_obs_cliente, text="Nenhum dado compatível.", text_color="gray").pack(pady=20)
        self.selecionados_cliente.intersection_update(valid_clientes)

        for widget in self.scroll_obs_adic.winfo_children(): widget.destroy()
        self.checkboxes_obs_adic.clear()
        if valid_adics:
            for obs in sorted(list(valid_adics)):
                cb = ctk.CTkCheckBox(self.scroll_obs_adic, text=obs, font=("Segoe UI", 12), command=self.cb_adic_clicked)
                cb.pack(anchor="w", padx=10, pady=5)
                if obs in self.selecionados_adic: cb.select()
                self.checkboxes_obs_adic.append(cb)
        else:
            ctk.CTkLabel(self.scroll_obs_adic, text="Nenhum dado compatível.", text_color="gray").pack(pady=20)
        self.selecionados_adic.intersection_update(valid_adics)

    # Métodos novos para o Controller manipular
    def get_filtros_selecionados(self):
        obs_cliente = [cb.cget("text") for cb in self.checkboxes_obs_cliente if cb.get()]
        obs_adic = [cb.cget("text") for cb in self.checkboxes_obs_adic if cb.get()]
        return obs_cliente, obs_adic