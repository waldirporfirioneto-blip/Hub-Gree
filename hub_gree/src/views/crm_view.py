import customtkinter as ctk
from tkinter import filedialog

class CRMView(ctk.CTkFrame):
    def __init__(self, master, controller):
        super().__init__(master, fg_color="transparent")
        self.controller = controller # Injeção de dependência do Controller
        self.criar_interface()

    def criar_interface(self):
        # UI encapsulada sem lógica de negócio
        self.btnArquivo = ctk.CTkButton(
            self, text="Selecionar Arquivos", command=self.solicitar_arquivos
        )
        self.btnArquivo.pack(pady=20)
        
        self.lbl_status = ctk.CTkLabel(self, text="Aguardando...")
        self.lbl_status.pack()

    def solicitar_arquivos(self):
        caminhos = filedialog.askopenfilenames(
            title="Selecione as planilhas", 
            filetypes=[("Arquivos Excel", "*.xlsx")]
        )
        if caminhos:
            self.lbl_status.configure(text=f"{len(caminhos)} arquivo(s) selecionado(s)")
            # Delega a responsabilidade de processamento ao Controller
            self.controller.processar_arquivos_lote(caminhos, self.atualizar_status)

    def atualizar_status(self, mensagem, concluido=False):
        """ Callback atualizado pelo Controller para refletir progresso na UI """
        self.lbl_status.configure(text=mensagem)
        if concluido:
            self.lbl_status.configure(text_color="green")