import customtkinter as ctk
from PIL import Image
from utils.config_app import CAMINHO_LOGO, CAMINHO_ICO

class HubLogisticaView(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("GREE - Hub de Ferramentas Logística")
        
        try:
            self.iconbitmap(CAMINHO_ICO)
        except Exception:
            try:
                from PIL import ImageTk
                logo_icone = ImageTk.PhotoImage(Image.open(CAMINHO_LOGO))
                self.iconphoto(False, logo_icone)
            except Exception:
                pass

        self.cor_sidebar = "#0A2540"
        self.cor_ativa = "#F26522"
        self.cor_hover = "#123150"
        self.cinza_fundo = "#F4F6F9"

        largura_app = 1180
        altura_app = 780
        largura_tela = self.winfo_screenwidth()
        altura_tela = self.winfo_screenheight()
        pos_x = (largura_tela // 2) - (largura_app // 2)
        pos_y = (altura_tela // 2) - (altura_app // 2)

        self.geometry(f"{largura_app}x{altura_app}+{pos_x}+{pos_y}")

        self.sidebar = None
        self.content_frame = None
        
        self.construir_layout_principal()

    def construir_layout_principal(self):
        #Barra Lateral
        self.sidebar = ctk.CTkFrame(self, width=260, corner_radius=0, fg_color=self.cor_sidebar)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        try:
            logo_pil = Image.open(CAMINHO_LOGO)
            logo_img = ctk.CTkImage(light_image=logo_pil, dark_image=logo_pil, size=(170, 40))
            self.lbl_logo = ctk.CTkLabel(self.sidebar, image=logo_img, text="")
            self.lbl_logo.pack(pady=(40, 20))
        except Exception:
            self.lbl_logo = ctk.CTkLabel(self.sidebar, text="GREE LOGÍSTICA", font=("Segoe UI", 20, "bold"), text_color="white")
            self.lbl_logo.pack(pady=(45, 20))

        divisoria = ctk.CTkFrame(self.sidebar, height=1, fg_color="#1C3F60")
        divisoria.pack(fill="x", padx=25, pady=(0, 25))

        font_menu = ("Segoe UI", 14, "bold")
        estilo_btn = {
            "fg_color": "transparent",
            "text_color": "#B0C4DE",
            "hover_color": self.cor_hover,
            "anchor": "w",
            "height": 45,
            "corner_radius": 8,
            "font": font_menu
        }

        self.btn_crm_side = ctk.CTkButton(self.sidebar, text="  📊   CRM Formatação", **estilo_btn)
        self.btn_crm_side.pack(fill="x", padx=15, pady=5)

        self.btn_cot_side = ctk.CTkButton(self.sidebar, text="  💰   Gerador de Cotações", **estilo_btn)
        self.btn_cot_side.pack(fill="x", padx=15, pady=5)

        self.btn_consolidador_side = ctk.CTkButton(self.sidebar, text="  📦   Consolidador PDFs", **estilo_btn)
        self.btn_consolidador_side.pack(fill="x", padx=15, pady=5)

        texto_assinatura = "Desenvolvido por:\nWaldir Neto & Vinicius Diogo\n\n\"Movendo o futuro da Gree\"\n推动格力的未来"
        self.lbl_desenvolvedores = ctk.CTkLabel(self.sidebar, text=texto_assinatura, font=("Segoe UI", 11), text_color="#5D7A99", justify="center")
        self.lbl_desenvolvedores.pack(side="bottom", pady=30)
        self.content_frame = ctk.CTkFrame(self, fg_color=self.cinza_fundo, corner_radius=0)
        self.content_frame.pack(side="right", fill="both", expand=True)

    def destacar_botao_ativo(self, botao_ativo):
        # Reseta as cores
        for btn in [self.btn_crm_side, self.btn_cot_side, self.btn_consolidador_side]:
            btn.configure(fg_color="transparent", text_color="#B0C4DE", hover_color=self.cor_hover)
        # Ativa a cor GREE no botão selecionado
        botao_ativo.configure(fg_color=self.cor_ativa, text_color="white", hover_color="#D85A1E")