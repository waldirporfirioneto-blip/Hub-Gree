import customtkinter as ctk
import tkinter as tk
from tkinter import ttk
from PIL import Image
from utils.config_app import CAMINHO_LOGO

class CotacaoView(ctk.CTkFrame):
    def __init__(self, master=None):
        super().__init__(master, fg_color="transparent")
        self.azul_gree = "#0057B8"
        self.laranja_gree = "#F26522"
        self.verde_gree = "#2EAF4A"
        self.setup_ui()

    def setup_ui(self):
        header = ctk.CTkFrame(self, fg_color=self.azul_gree, corner_radius=0)
        header.pack(fill="x", side="top")

        try:
            logo = ctk.CTkImage(light_image=Image.open(CAMINHO_LOGO), size=(130, 30))
            lblLogo = ctk.CTkLabel(header, image=logo, text="")
            lblLogo.pack(side="left", padx=25, pady=12)
        except Exception:
            pass

        title = ctk.CTkLabel(header, text="Gerador de Cotações", text_color="white", font=("Segoe UI", 18, "bold"))
        title.pack(side="left", padx=15, pady=12)

        body = ctk.CTkFrame(self, fg_color="white", corner_radius=15)
        body.pack(fill="both", expand=True, padx=25, pady=20)

        lbl1 = ctk.CTkLabel(body, text=" 📂  1. Importe a planilha de dados", font=("Segoe UI", 17, "bold"), text_color="#1C2D42")
        lbl1.pack(anchor="w", padx=30, pady=(20, 8))

        self.btn_carregar = ctk.CTkButton(
            body, width=220, height=40, corner_radius=10, 
            fg_color=self.laranja_gree, hover_color="#D35400", 
            text="Importar Base de Dados", font=("Segoe UI", 13, "bold")
        )
        self.btn_carregar.pack(padx=30, fill="x")

        self.lbl_status = tk.Label(body, text="Nenhum arquivo carregado", bg="white", fg="#7F8C8D", font=("Segoe UI", 12, "bold"))
        self.lbl_status.pack(anchor="w", padx=35, pady=5)

        lbl2 = ctk.CTkLabel(body, text=" 📅  2. Organize a ordem das Cotações", font=("Segoe UI", 17, "bold"), text_color="#1C2D42")
        lbl2.pack(anchor="w", padx=30, pady=(15, 8))

        lbl_instrucao = ctk.CTkLabel(body, text="Arraste os itens na lista ou utilize os botões de subir e descer ao lado para ordenar:", font=("Segoe UI", 12, "italic"), text_color="gray50")
        lbl_instrucao.pack(anchor="w", padx=35, pady=(0, 5))

        frame_meio = ctk.CTkFrame(body, fg_color="transparent")
        frame_meio.pack(fill="both", expand=True, padx=30, pady=5)

        frame_lista = ctk.CTkFrame(frame_meio, fg_color="white", border_color="#BDC3C7", border_width=1, corner_radius=8)
        frame_lista.pack(side="left", fill="both", expand=True)

        self.listbox = tk.Listbox(
            frame_lista, font=("Segoe UI", 12), selectbackground=self.azul_gree, 
            selectforeground="white", activestyle="none", relief="flat", 
            borderwidth=0, highlightthickness=0, bg="white", fg="#2C3E50"
        )
        self.listbox.pack(side="left", fill="both", expand=True, padx=8, pady=8)

        scrollbar = ttk.Scrollbar(frame_lista, orient="vertical", command=self.listbox.yview)
        scrollbar.pack(side="right", fill="y", padx=(0, 4), pady=4)
        self.listbox.config(yscrollcommand=scrollbar.set)

        frame_botoes = ctk.CTkFrame(frame_meio, fg_color="transparent")
        frame_botoes.pack(side="right", fill="y", padx=(15, 0))

        btn_up = ctk.CTkButton(frame_botoes, text="▲ Subir", command=self.mover_cima, fg_color="#EAF2FC", text_color=self.azul_gree, hover_color="#D0E1F9", font=("Segoe UI", 12, "bold"), width=110, height=40, corner_radius=8)
        btn_up.pack(pady=(20, 10))

        btn_down = ctk.CTkButton(frame_botoes, text="▼ Descer", command=self.mover_baixo, fg_color="#EAF2FC", text_color=self.azul_gree, hover_color="#D0E1F9", font=("Segoe UI", 12, "bold"), width=110, height=40, corner_radius=8)
        btn_down.pack(pady=10)

        self.listbox.bind("<Button-1>", self.on_drag_start)
        self.listbox.bind("<B1-Motion>", self.on_drag_motion)

        lbl3 = ctk.CTkLabel(body, text=" 📥  3. Gere a cotação final formatada", font=("Segoe UI", 17, "bold"), text_color="#1C2D42")
        lbl3.pack(anchor="w", padx=30, pady=(15, 8))

        self.btn_gerar = ctk.CTkButton(
            body, width=260, height=55, corner_radius=10, 
            fg_color="#BDC3C7", text_color="white", text="EXPORTAR COTAÇÃO FORMATADA", 
            font=("Segoe UI", 16, "bold"), state="disabled"
        )
        self.btn_gerar.pack(padx=30, fill="x", pady=(0, 20))

        footer = ctk.CTkLabel(self, text="Desenvolvido para Operação GREE LOGÍSTICA", font=("Segoe UI", 10, "italic"), text_color="gray50")
        footer.pack(pady=8)

    # Lógica puramente visual e de interface (Mantida na View)
    def atualizar_status(self, texto, cor):
        self.lbl_status.config(text=texto, fg=cor)

    def preencher_lista(self, itens):
        self.listbox.delete(0, tk.END)
        for item in itens:
            self.listbox.insert(tk.END, item)

    def obter_ordem_lista(self):
        return list(self.listbox.get(0, tk.END))

    def ativar_botao_gerar(self):
        self.btn_gerar.configure(state="normal", fg_color=self.verde_gree, hover_color="#22953D")

    def desativar_botao_gerar(self):
        self.btn_gerar.configure(state="disabled", fg_color="#BDC3C7")

    def mover_cima(self):
        try:
            sel = self.listbox.curselection()[0]
            if sel > 0:
                text = self.listbox.get(sel)
                self.listbox.delete(sel)
                self.listbox.insert(sel - 1, text)
                self.listbox.select_set(sel - 1)
        except IndexError: pass

    def mover_baixo(self):
        try:
            sel = self.listbox.curselection()[0]
            if sel < self.listbox.size() - 1:
                text = self.listbox.get(sel)
                self.listbox.delete(sel)
                self.listbox.insert(sel + 1, text)
                self.listbox.select_set(sel + 1)
        except IndexError: pass

    def on_drag_start(self, event):
        self._drag_start_index = self.listbox.nearest(event.y)

    def on_drag_motion(self, event):
        i = self.listbox.nearest(event.y)
        if i < self.listbox.size() and i >= 0:
            if i != self._drag_start_index:
                item = self.listbox.get(self._drag_start_index)
                self.listbox.delete(self._drag_start_index)
                self.listbox.insert(i, item)
                self.listbox.select_clear(0, tk.END)
                self.listbox.select_set(i)
                self._drag_start_index = i