import customtkinter as ctk
from src.controllers.crm_controller import CRMController
from src.views.crm_view import CRMView

def main():
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")
    
    app = ctk.CTk()
    app.title("GREE - Hub de Ferramentas Logística")
    app.geometry("800x600")
    
    # Injeção de Dependências
    crm_controller = CRMController()
    crm_view = CRMView(app, crm_controller)
    crm_view.pack(fill="both", expand=True)
    
    app.mainloop()

if __name__ == "__main__":
    main()