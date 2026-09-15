import customtkinter as ctk
from views.main_window import HubLogisticaView
from controllers.cotacao_controller import CotacaoController
from controllers.consolidador_controller import ConsolidadorController
from controllers.crm_controller import CRMController

class MainController:
    def __init__(self):
        
        self.view = HubLogisticaView()
        self.ferramenta_atual = None
        self.view.btn_crm_side.configure(command=lambda: self.abrir_modulo("CRM"))
        self.view.btn_cot_side.configure(command=lambda: self.abrir_modulo("COTACAO"))
        self.view.btn_consolidador_side.configure(command=lambda: self.abrir_modulo("CONSOLIDADOR"))
        self.abrir_modulo("CRM")

    def limpar_tela_conteudo(self):
        if self.ferramenta_atual:
            self.ferramenta_atual.destroy()
            

    def abrir_modulo(self, nome_modulo):
        self.limpar_tela_conteudo()

        if nome_modulo == "CRM":
            self.view.destacar_botao_ativo(self.view.btn_crm_side)
            controlador = CRMController(self.view.content_frame)
            self.ferramenta_atual = controlador.get_view()
            self.ferramenta_atual.pack(fill="both", expand=True)
            
        elif nome_modulo == "COTACAO":
            self.view.destacar_botao_ativo(self.view.btn_cot_side)
            controlador = CotacaoController(self.view.content_frame)
            self.ferramenta_atual = controlador.get_view()
            self.ferramenta_atual.pack(fill="both", expand=True)
            
        elif nome_modulo == "CONSOLIDADOR":
            self.view.destacar_botao_ativo(self.view.btn_consolidador_side)
            controlador = ConsolidadorController(self.view.content_frame)
            self.ferramenta_atual = controlador.get_view()
            self.ferramenta_atual.pack(fill="both", expand=True)
            
    def criar_tela_aviso(self, mensagem):
        self.ferramenta_atual = ctk.CTkFrame(self.view.content_frame, fg_color="transparent")
        self.ferramenta_atual.pack(fill="both", expand=True)
        
        lbl = ctk.CTkLabel(
            self.ferramenta_atual, 
            text=mensagem, 
            font=("Segoe UI", 24, "bold"), 
            text_color="gray"
        )
        lbl.pack(expand=True)

    def iniciar(self):
        self.view.mainloop()