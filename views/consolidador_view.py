import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image
import queue
from datetime import datetime as dt
from utils.config_app import CAMINHO_LOGO
from utils.config_manager import ConfigManager

class ConsolidadorView(ctk.CTkFrame):
    def __init__(self, master=None):
        super().__init__(master, fg_color="transparent")
        self.queue = queue.Queue()
        self.arquivos_mescla_lista = []
        self.arquivos_jpg_lista = []
        
        self.definir_estilos()
        self.criar_widgets()
        self.verificar_queue()

    def definir_estilos(self):
        style = ttk.Style()
        style.theme_use("clam")
        azul_gree = "#0057B8"
        style.configure(".", background="white", font=("Segoe UI", 10))
        style.configure("TNotebook", background="white", borderwidth=0)
        style.configure("TNotebook.Tab", background="#EAF2FC", foreground="#2C3E50", padding=[2, 3], font=("Segoe UI", 10, "bold"))
        style.map("TNotebook.Tab", background=[("selected", azul_gree)], foreground=[("selected", "#FFFFFF")])
        style.configure("TLabelframe", background="#FFFFFF", bordercolor="#E0E6ED", borderwidth=1)
        style.configure("TLabelframe.Label", font=("Segoe UI", 11, "bold"), foreground=azul_gree, background="#FFFFFF")
        style.configure("TLabel", background="#FFFFFF", foreground="#2C3E50")
        style.configure("TProgressbar", thickness=8, troughcolor="#E0E6ED", background=azul_gree)

    def criar_widgets(self):
        header_frame = tk.Frame(self, bg="#0057B8", height=70)
        header_frame.pack(fill="x", side="top")
        header_frame.pack_propagate(False)

        try:
            logo = ctk.CTkImage(light_image=Image.open(CAMINHO_LOGO), size=(110, 30))
            lblLogo = ctk.CTkLabel(header_frame, image=logo, text="")
            lblLogo.pack(side="left", padx=25, pady=17)
        except: pass

        title_text_frame = tk.Frame(header_frame, bg="#0057B8")
        title_text_frame.pack(side="left", fill="both", expand=True, pady=10)
        tk.Label(title_text_frame, text="Gree Electric Appliances", font=("Segoe UI", 12, "bold"), fg="white", bg="#0057B8").pack(anchor="w", padx=(5, 15))
        tk.Label(title_text_frame, text="Painel Integrado de Automação e Gerenciamento de Documentos", font=("Segoe UI", 9, "italic"), fg="#D0E1F9", bg="#0057B8").pack(anchor="w", padx=(5, 15))

        body = ctk.CTkFrame(self, fg_color="white", corner_radius=15)
        body.pack(fill="both", expand=True, padx=25, pady=(20, 10))

        self.notebook = ttk.Notebook(body)
        self.notebook.pack(fill="both", expand=True, padx=15, pady=15)

        self.tab_vendas = ctk.CTkFrame(self.notebook, fg_color="white", corner_radius=0)
        self.tab_mesclar = ctk.CTkFrame(self.notebook, fg_color="white", corner_radius=0)
        self.tab_conversor = ctk.CTkFrame(self.notebook, fg_color="white", corner_radius=0)

        self.notebook.add(self.tab_vendas, text=" 📊  Relatório de Ordens de Venda")
        self.notebook.add(self.tab_mesclar, text=" 🔗  Mesclador de PDFs")
        self.notebook.add(self.tab_conversor, text=" ⚙️  Conversor de Arquivos")

        self.montar_aba_vendas()
        self.montar_aba_mesclar()
        self.montar_aba_conversor()

        status_frame = ctk.CTkFrame(self, fg_color="white", corner_radius=12, border_color="#E0E6ED", border_width=1)
        status_frame.pack(fill="x", side="bottom", padx=25, pady=(0, 20), before=body)

        progress_container = ctk.CTkFrame(status_frame, fg_color="transparent")
        progress_container.pack(fill="x", padx=15, pady=(15, 8))

        self.progress = ctk.CTkProgressBar(progress_container, height=14, progress_color="#2EAF4A", fg_color="#E0E6ED")
        self.progress.pack(side="left", fill="x", expand=True, padx=(0, 15))
        self.progress.set(0)

        self.lbl_porcentagem = ctk.CTkLabel(progress_container, text="0%", font=("Segoe UI", 14, "bold"), text_color="#0057B8")
        self.lbl_porcentagem.pack(side="right")

        self.log_text = tk.Text(status_frame, height=5, font=("Consolas", 9), bg="#FFFFFF", fg="#2C3E50", insertbackground="#0057B8", relief="flat", borderwidth=0, highlightthickness=0)
        self.log_text.pack(fill="x", padx=15, pady=(0, 15))
        self.log_text.configure(state="disabled")

        self.escrever_log("Sistema inicializado com sucesso. Aguardando comandos...\nDica: Para evitar travamentos, as operações rodam em segundo plano.")

    def escrever_log(self, mensagem):
        timestamp = dt.now().strftime("[%H:%M:%S]")
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"{timestamp} {mensagem}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def verificar_queue(self):
        try:
            while True:
                msg_type, payload = self.queue.get_nowait()
                if msg_type == "log": self.escrever_log(payload)
                elif msg_type == "progress":
                    self.progress.set(payload / 100.0)
                    self.lbl_porcentagem.configure(text=f"{int(payload)}%")
                elif msg_type == "success":
                    self.escrever_log(payload)
                    messagebox.showinfo("Sucesso", payload)
                    self.progress.set(1.0)
                    self.lbl_porcentagem.configure(text="100%")
                elif msg_type == "error":
                    self.escrever_log(f"❌ {payload}")
                    messagebox.showerror("Erro", payload)
                    self.progress.set(0)
                    self.lbl_porcentagem.configure(text="0%")
                self.queue.task_done()
        except queue.Empty: pass
        finally: self.after(100, self.verificar_queue)

    def montar_aba_vendas(self):
        container = ctk.CTkFrame(self.tab_vendas, fg_color="white", corner_radius=0)
        container.pack(fill="both", expand=True, padx=15, pady=15)
        
        mode_group = ctk.CTkFrame(container, fg_color="white", border_color="#E0E6ED", border_width=1, corner_radius=8)
        mode_group.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(mode_group, text="1. Modo de Leitura", font=("Segoe UI", 11, "bold"), text_color="#0057B8").pack(anchor="w", padx=10, pady=(8, 2))
        inner_mode = ctk.CTkFrame(mode_group, fg_color="transparent")
        inner_mode.pack(fill="x", padx=10, pady=(0, 10))

        self.modo_leitura_var = tk.StringVar(value="pasta")
        r1 = ctk.CTkRadioButton(inner_mode, text="Pasta de PDFs", variable=self.modo_leitura_var, value="pasta", fg_color="#0057B8", font=("Segoe UI", 11), text_color="#2C3E50", command=self.atualizar_label_selecao)
        r1.pack(side="left", padx=(0, 20), pady=8)
        r2 = ctk.CTkRadioButton(inner_mode, text="Arquivo PDF Único", variable=self.modo_leitura_var, value="arquivo", fg_color="#0057B8", font=("Segoe UI", 11), text_color="#2C3E50", command=self.atualizar_label_selecao)
        r2.pack(side="left", pady=8)

        selection_group = ctk.CTkFrame(container, fg_color="white", border_color="#E0E6ED", border_width=1, corner_radius=8)
        selection_group.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(selection_group, text="2. Seleção de Entrada", font=("Segoe UI", 11, "bold"), text_color="#0057B8").pack(anchor="w", padx=10, pady=(8, 2))
        inner_selection = ctk.CTkFrame(selection_group, fg_color="transparent")
        inner_selection.pack(fill="x", padx=10, pady=(0, 10))
        inner_selection.columnconfigure(1, weight=1)

        self.lbl_entrada_tipo = ttk.Label(inner_selection, text="Pasta de PDFs:")
        self.lbl_entrada_tipo.grid(row=0, column=0, sticky="w", padx=(0, 5))
        self.ent_vendas_path = ctk.CTkEntry(inner_selection, height=32, corner_radius=6)
        self.ent_vendas_path.grid(row=0, column=1, sticky="ew", padx=(0, 5))

        self.btn_vendas_proc = ctk.CTkButton(inner_selection, text="Procurar...", width=100, height=32, corner_radius=6, fg_color="#F26522", hover_color="#D35400", text_color="white", font=("Segoe UI", 11, "bold"))
        self.btn_vendas_proc.grid(row=0, column=2, sticky="e")

        config_group = ctk.CTkFrame(container, fg_color="white", border_color="#E0E6ED", border_width=1, corner_radius=8)
        config_group.pack(fill="x", pady=(0, 15))
        ctk.CTkLabel(config_group, text="3. Configurações do Relatório", font=("Segoe UI", 11, "bold"), text_color="#0057B8").pack(anchor="w", padx=10, pady=(8, 2))
        inner_config = ctk.CTkFrame(config_group, fg_color="transparent")
        inner_config.pack(fill="x", padx=10, pady=(0, 10))

        tk.Label(inner_config, text="Data Doc:", bg="#FFFFFF", font=("Segoe UI", 10)).grid(row=0, column=0, sticky="w", padx=(0, 5), pady=5)
        self.ent_data_doc = ctk.CTkEntry(inner_config, width=120, height=32, corner_radius=6)
        self.ent_data_doc.insert(0, dt.now().strftime("%d/%m/%Y"))
        self.ent_data_doc.grid(row=0, column=1, sticky="w", padx=(0, 20), pady=5)

        tk.Label(inner_config, text="Solicitação:", bg="#FFFFFF", font=("Segoe UI", 10)).grid(row=0, column=2, sticky="w", padx=(0, 5), pady=5)
        
        self.ent_solicitacao = ctk.CTkEntry(inner_config, width=140, height=32, corner_radius=6)
        
        # --- BUSCA A INFORMAÇÃO DO JSON AQUI ---
        cfg = ConfigManager()
        self.ent_solicitacao.insert(0, cfg.get("solicitacao_padrao_consolidador"))
        
        self.ent_solicitacao.grid(row=0, column=3, sticky="w", padx=(0, 20), pady=5)

        self.btn_vendas_gerar = ctk.CTkButton(container, text="AUTOMATIZAR E GERAR EXCEL", height=45, corner_radius=8, fg_color="#0057B8", hover_color="#003A6F", font=("Segoe UI", 13, "bold"))
        self.btn_vendas_gerar.pack(fill="x", pady=(5, 0))

    def atualizar_label_selecao(self):
        modo = self.modo_leitura_var.get()
        if modo == "pasta": self.lbl_entrada_tipo.configure(text="Pasta de PDFs:")
        else: self.lbl_entrada_tipo.configure(text="Arquivo PDF:")

    def montar_aba_mesclar(self):
        container = ctk.CTkFrame(self.tab_mesclar, fg_color="white", corner_radius=0)
        container.pack(fill="both", expand=True, padx=15, pady=15)
        ctk.CTkLabel(container, text="Junte múltiplos PDFs de faturamento, Notas Fiscais ou laudos em um único arquivo consolidado.", font=("Segoe UI", 11, "italic"), text_color="gray40").pack(anchor="w", pady=(0, 10))

        list_frame = ctk.CTkFrame(container, fg_color="transparent")
        list_frame.pack(fill="both", expand=True, pady=(0, 10))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        frame_lista = ctk.CTkFrame(list_frame, fg_color="white", border_color="#BDC3C7", border_width=1, corner_radius=8)
        frame_lista.grid(row=0, column=0, sticky="nsew")

        self.lst_pdfs = tk.Listbox(frame_lista, font=("Segoe UI", 11), selectbackground="#0057B8", selectforeground="white", relief="flat", borderwidth=0, highlightthickness=0, bg="white", fg="#2C3E50")
        self.lst_pdfs.pack(side="left", fill="both", expand=True, padx=8, pady=8)
        scrollbar = ttk.Scrollbar(frame_lista, orient="vertical", command=self.lst_pdfs.yview)
        scrollbar.pack(side="right", fill="y", padx=(0, 4), pady=4)
        self.lst_pdfs.configure(yscrollcommand=scrollbar.set)

        btn_control_frame = ctk.CTkFrame(list_frame, fg_color="transparent")
        btn_control_frame.grid(row=0, column=2, sticky="ns", padx=(10, 0))

        self.btn_mescla_up = ctk.CTkButton(btn_control_frame, text="Mover ↑", command=self.mover_item_cima, width=100, height=32, corner_radius=6, fg_color="#0057B8", text_color="white", hover_color="#003A6F", font=("Segoe UI", 10, "bold"))
        self.btn_mescla_up.pack(fill="x", pady=2)
        self.btn_mescla_down = ctk.CTkButton(btn_control_frame, text="Mover ↓", command=self.mover_item_baixo, width=100, height=32, corner_radius=6, fg_color="#0057B8", text_color="white", hover_color="#003A6F", font=("Segoe UI", 10, "bold"))
        self.btn_mescla_down.pack(fill="x", pady=2)
        self.btn_mescla_rem = ctk.CTkButton(btn_control_frame, text="Remover", command=self.remover_item, width=100, height=32, corner_radius=6, fg_color="#E74C3C", text_color="white", hover_color="#C0392B", font=("Segoe UI", 10, "bold"))
        self.btn_mescla_rem.pack(fill="x", pady=15)
        self.btn_mescla_clean = ctk.CTkButton(btn_control_frame, text="Limpar Lista", command=self.limpar_lista, width=100, height=32, corner_radius=6, fg_color="#7F8C8D", text_color="white", hover_color="#626567", font=("Segoe UI", 10, "bold"))
        self.btn_mescla_clean.pack(fill="x", pady=2)

        action_bar = ctk.CTkFrame(container, fg_color="transparent")
        action_bar.pack(fill="x", side="bottom", pady=(10, 0))
        action_bar.columnconfigure(2, weight=1)

        self.btn_add_files = ctk.CTkButton(action_bar, text="+ Adicionar PDFs individuais...", height=36, corner_radius=8, fg_color="#F26522", text_color="white", hover_color="#D35400", font=("Segoe UI", 11, "bold"))
        self.btn_add_files.grid(row=0, column=0, sticky="w", padx=(0, 10))
        self.btn_add_dir = ctk.CTkButton(action_bar, text="+ Importar de uma pasta...", height=36, corner_radius=8, fg_color="#F26522", text_color="white", hover_color="#D35400", font=("Segoe UI", 11, "bold"))
        self.btn_add_dir.grid(row=0, column=1, sticky="w")
        self.btn_consolidar = ctk.CTkButton(action_bar, text="MESCLAR E SALVAR PDF ÚNICO", height=36, width=450, corner_radius=8, fg_color="#2EAF4A", hover_color="#22953D", font=("Segoe UI", 12, "bold"))
        self.btn_consolidar.grid(row=0, column=3, sticky="e")

    def mover_item_cima(self):
        sel = self.lst_pdfs.curselection()
        if not sel: return
        idx = sel[0]
        if idx == 0: return
        self.arquivos_mescla_lista[idx], self.arquivos_mescla_lista[idx - 1] = self.arquivos_mescla_lista[idx - 1], self.arquivos_mescla_lista[idx]
        text = self.lst_pdfs.get(idx)
        self.lst_pdfs.delete(idx)
        self.lst_pdfs.insert(idx - 1, text)
        self.lst_pdfs.selection_set(idx - 1)

    def mover_item_baixo(self):
        sel = self.lst_pdfs.curselection()
        if not sel: return
        idx = sel[0]
        if idx == self.lst_pdfs.size() - 1: return
        self.arquivos_mescla_lista[idx], self.arquivos_mescla_lista[idx + 1] = self.arquivos_mescla_lista[idx + 1], self.arquivos_mescla_lista[idx]
        text = self.lst_pdfs.get(idx)
        self.lst_pdfs.delete(idx)
        self.lst_pdfs.insert(idx + 1, text)
        self.lst_pdfs.selection_set(idx + 1)

    def remover_item(self):
        sel = self.lst_pdfs.curselection()
        if not sel: return
        idx = sel[0]
        self.lst_pdfs.delete(idx)
        self.arquivos_mescla_lista.pop(idx)

    def limpar_lista(self):
        self.lst_pdfs.delete(0, "end")
        self.arquivos_mescla_lista.clear()

    def montar_aba_conversor(self):
        container = ctk.CTkFrame(self.tab_conversor, fg_color="white", corner_radius=0)
        container.pack(fill="both", expand=True, padx=15, pady=15)
        ctk.CTkLabel(container, text="Selecione abaixo a ferramenta de conversão rápida que deseja utilizar.", font=("Segoe UI", 11, "italic"), text_color="gray40").pack(anchor="w", pady=(0, 10))

        self.sub_notebook = ttk.Notebook(container)
        self.sub_notebook.pack(fill="both", expand=True)
        self.tab_pdf_word = ctk.CTkFrame(self.sub_notebook, fg_color="white", corner_radius=0)
        self.tab_pdf_jpg = ctk.CTkFrame(self.sub_notebook, fg_color="white", corner_radius=0)
        self.tab_jpg_pdf = ctk.CTkFrame(self.sub_notebook, fg_color="white", corner_radius=0)

        self.sub_notebook.add(self.tab_pdf_word, text="PDF para Word")
        self.sub_notebook.add(self.tab_pdf_jpg, text="PDF para JPG")
        self.sub_notebook.add(self.tab_jpg_pdf, text="JPG para PDF")

        self.montar_sub_pdf_word()
        self.montar_sub_pdf_jpg()
        self.montar_sub_jpg_pdf()

    def montar_sub_pdf_word(self):
        container = ctk.CTkFrame(self.tab_pdf_word, fg_color="white", corner_radius=0)
        container.pack(fill="both", expand=True, padx=15, pady=15)
        ctk.CTkLabel(container, text="Converta arquivos PDF em documentos do Word (.docx) editáveis.", font=("Segoe UI", 11, "italic"), text_color="gray40").pack(anchor="w", pady=(0, 15))

        selection_group = ctk.CTkFrame(container, fg_color="white", border_color="#E0E6ED", border_width=1, corner_radius=8)
        selection_group.pack(fill="x", pady=(0, 20))
        ctk.CTkLabel(selection_group, text="Seleção de Arquivo", font=("Segoe UI", 11, "bold"), text_color="#0057B8").pack(anchor="w", padx=10, pady=(8, 2))
        
        inner_selection = ctk.CTkFrame(selection_group, fg_color="transparent")
        inner_selection.pack(fill="x", padx=10, pady=(0, 10))
        inner_selection.columnconfigure(1, weight=1)

        ctk.CTkLabel(inner_selection, text="Arquivo PDF:", font=("Segoe UI", 11), text_color="#2C3E50").grid(row=0, column=0, sticky="w", padx=(0, 5))
        self.ent_pdf_word_path = ctk.CTkEntry(inner_selection, height=32, corner_radius=6)
        self.ent_pdf_word_path.grid(row=0, column=1, sticky="ew", padx=(0, 5))

        self.btn_proc_word = ctk.CTkButton(inner_selection, text="Procurar...", width=100, height=32, corner_radius=6, fg_color="#F26522", hover_color="#D35400", text_color="white", font=("Segoe UI", 11, "bold"))
        self.btn_proc_word.grid(row=0, column=2, sticky="e")

        self.btn_gerar_word = ctk.CTkButton(container, text="CONVERTER PARA WORD", height=45, corner_radius=8, fg_color="#0057B8", hover_color="#003A6F", font=("Segoe UI", 13, "bold"))
        self.btn_gerar_word.pack(fill="x", pady=(5, 0))

    def montar_sub_pdf_jpg(self):
        container = ctk.CTkFrame(self.tab_pdf_jpg, fg_color="white", corner_radius=0)
        container.pack(fill="both", expand=True, padx=15, pady=15)
        ctk.CTkLabel(container, text="Extraia as páginas do PDF para imagens JPG.", font=("Segoe UI", 11, "italic"), text_color="gray40").pack(anchor="w", pady=(0, 15))

        config_group = ctk.CTkFrame(container, fg_color="white", border_color="#E0E6ED", border_width=1, corner_radius=8)
        config_group.pack(fill="x", pady=(0, 20))
        ctk.CTkLabel(config_group, text="Seleção de Arquivo e Opções", font=("Segoe UI", 11, "bold"), text_color="#0057B8").pack(anchor="w", padx=10, pady=(8, 2))

        inner_config = ctk.CTkFrame(config_group, fg_color="transparent")
        inner_config.pack(fill="x", padx=10, pady=(0, 10))
        inner_config.columnconfigure(1, weight=1)

        ctk.CTkLabel(inner_config, text="Arquivo PDF:", font=("Segoe UI", 11), text_color="#2C3E50").grid(row=0, column=0, sticky="w", padx=(0, 5), pady=5)
        self.ent_pdf_jpg_path = ctk.CTkEntry(inner_config, height=32, corner_radius=6)
        self.ent_pdf_jpg_path.grid(row=0, column=1, sticky="ew", padx=(0, 5), pady=5)

        self.btn_proc_jpg = ctk.CTkButton(inner_config, text="Procurar...", width=100, height=32, corner_radius=6, fg_color="#F26522", hover_color="#D35400", text_color="white", font=("Segoe UI", 11, "bold"))
        self.btn_proc_jpg.grid(row=0, column=2, sticky="e", pady=5)

        ctk.CTkLabel(inner_config, text="Qualidade (DPI):", font=("Segoe UI", 11), text_color="#2C3E50").grid(row=1, column=0, sticky="w", padx=(0, 5), pady=5)
        
        cfg = ConfigManager()
        self.cmb_dpi = ttk.Combobox(inner_config, values=["150 (Médio)", "200 (Bom)", "300 (Alto)"], state="readonly", width=15)
        self.cmb_dpi.set(cfg.get("dpi_padrao_conversao"))
        self.cmb_dpi.grid(row=1, column=1, sticky="w", pady=5)

        self.btn_gerar_jpg = ctk.CTkButton(container, text="CONVERTER PARA JPG", height=45, corner_radius=8, fg_color="#2EAF4A", hover_color="#22953D", font=("Segoe UI", 13, "bold"))
        self.btn_gerar_jpg.pack(fill="x", pady=(5, 0))

    def montar_sub_jpg_pdf(self):
        container = ctk.CTkFrame(self.tab_jpg_pdf, fg_color="white", corner_radius=0)
        container.pack(fill="both", expand=True, padx=15, pady=10)
        ctk.CTkLabel(container, text="Una fotos em um único arquivo PDF.", font=("Segoe UI", 11, "italic"), text_color="gray40").pack(anchor="w", pady=(0, 5))

        config_group = ctk.CTkFrame(container, fg_color="white", border_color="#E0E6ED", border_width=1, corner_radius=8)
        config_group.pack(side="bottom", fill="x", pady=(0, 0))
        ctk.CTkLabel(config_group, text="Configurações de Layout", font=("Segoe UI", 11, "bold"), text_color="#0057B8").pack(anchor="w", padx=10, pady=(8, 2))

        inner_config = ctk.CTkFrame(config_group, fg_color="transparent")
        inner_config.pack(fill="x", padx=10, pady=(0, 10))
        inner_config.columnconfigure(4, weight=1)

        ctk.CTkLabel(inner_config, text="Orientação:", font=("Segoe UI", 11), text_color="#2C3E50").grid(row=0, column=0, sticky="w", padx=(0, 5))
        self.cmb_orientacao = ttk.Combobox(inner_config, values=["Retrato (Vertical)", "Paisagem (Horizontal)"], state="readonly", width=18)
        self.cmb_orientacao.set("Retrato (Vertical)")
        self.cmb_orientacao.grid(row=0, column=1, sticky="w", padx=(0, 15))

        ctk.CTkLabel(inner_config, text="Margem:", font=("Segoe UI", 11), text_color="#2C3E50").grid(row=0, column=2, sticky="w", padx=(0, 5))
        self.cmb_margem = ttk.Combobox(inner_config, values=["Sem margem", "Margem fina", "Margem larga"], state="readonly", width=15)
        self.cmb_margem.set("Sem margem")
        self.cmb_margem.grid(row=0, column=3, sticky="w")

        self.btn_gerar_pdf = ctk.CTkButton(inner_config, text="GERAR PDF", height=36, width=200, corner_radius=8, fg_color="#2EAF4A", hover_color="#22953D", font=("Segoe UI", 12, "bold"))
        self.btn_gerar_pdf.grid(row=0, column=5, sticky="e", padx=(10, 0))

        list_frame = ctk.CTkFrame(container, fg_color="transparent")
        list_frame.pack(side="top", fill="both", expand=True, pady=(0, 15))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        frame_lista = ctk.CTkFrame(list_frame, fg_color="white", border_color="#BDC3C7", border_width=1, corner_radius=8)
        frame_lista.grid(row=0, column=0, sticky="nsew")

        self.lst_jpgs = tk.Listbox(frame_lista, font=("Segoe UI", 11), selectbackground="#0057B8", selectforeground="white", relief="flat", borderwidth=0, highlightthickness=0, bg="white", fg="#2C3E50")
        self.lst_jpgs.pack(side="left", fill="both", expand=True, padx=8, pady=8)
        scrollbar = ttk.Scrollbar(frame_lista, orient="vertical", command=self.lst_jpgs.yview)
        scrollbar.pack(side="right", fill="y", padx=(0, 4), pady=4)
        self.lst_jpgs.configure(yscrollcommand=scrollbar.set)

        btn_control_frame = ctk.CTkFrame(list_frame, fg_color="transparent")
        btn_control_frame.grid(row=0, column=2, sticky="ns", padx=(10, 0))

        self.btn_add_img = ctk.CTkButton(btn_control_frame, text="Adicionar...", width=100, height=25, corner_radius=6, fg_color="#F26522", text_color="white", hover_color="#D35400", font=("Segoe UI", 10, "bold"))
        self.btn_add_img.pack(fill="x", pady=2)
        self.btn_img_up = ctk.CTkButton(btn_control_frame, text="Mover ↑", command=self.mover_jpg_cima, width=100, height=20, corner_radius=6, fg_color="#0057B8", text_color="white", hover_color="#003A6F", font=("Segoe UI", 10, "bold"))
        self.btn_img_up.pack(fill="x", pady=2)
        self.btn_img_down = ctk.CTkButton(btn_control_frame, text="Mover ↓", command=self.mover_jpg_baixo, width=100, height=20, corner_radius=6, fg_color="#0057B8", text_color="white", hover_color="#003A6F", font=("Segoe UI", 10, "bold"))
        self.btn_img_down.pack(fill="x", pady=2)
        self.btn_img_rem = ctk.CTkButton(btn_control_frame, text="Remover", command=self.remover_jpg_item, width=100, height=20, corner_radius=6, fg_color="#E74C3C", text_color="white", hover_color="#C0392B", font=("Segoe UI", 10, "bold"))
        self.btn_img_rem.pack(fill="x", pady=2)
        self.btn_img_clean = ctk.CTkButton(btn_control_frame, text="Limpar", command=self.limpar_jpg_lista, width=100, height=20, corner_radius=6, fg_color="#7F8C8D", text_color="white", hover_color="#626567", font=("Segoe UI", 10, "bold"))
        self.btn_img_clean.pack(fill="x", pady=2)

    def mover_jpg_cima(self):
        sel = self.lst_jpgs.curselection()
        if not sel: return
        idx = sel[0]
        if idx == 0: return
        self.arquivos_jpg_lista[idx], self.arquivos_jpg_lista[idx - 1] = self.arquivos_jpg_lista[idx - 1], self.arquivos_jpg_lista[idx]
        text = self.lst_jpgs.get(idx)
        self.lst_jpgs.delete(idx)
        self.lst_jpgs.insert(idx - 1, text)
        self.lst_jpgs.selection_set(idx - 1)

    def mover_jpg_baixo(self):
        sel = self.lst_jpgs.curselection()
        if not sel: return
        idx = sel[0]
        if idx == self.lst_jpgs.size() - 1: return
        self.arquivos_jpg_lista[idx], self.arquivos_jpg_lista[idx + 1] = self.arquivos_jpg_lista[idx + 1], self.arquivos_jpg_lista[idx]
        text = self.lst_jpgs.get(idx)
        self.lst_jpgs.delete(idx)
        self.lst_jpgs.insert(idx + 1, text)
        self.lst_jpgs.selection_set(idx + 1)

    def remover_jpg_item(self):
        sel = self.lst_jpgs.curselection()
        if not sel: return
        idx = sel[0]
        self.lst_jpgs.delete(idx)
        self.arquivos_jpg_lista.pop(idx)

    def limpar_jpg_lista(self):
        self.lst_jpgs.delete(0, "end")
        self.arquivos_jpg_lista.clear()