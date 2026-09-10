import os
import threading
from tkinter import filedialog, messagebox
from views.consolidador_view import ConsolidadorView
from models.consolidador_model import ConsolidadorModel
from datetime import datetime as dt

class ConsolidadorController:
    def __init__(self, parent_frame):
        self.model = ConsolidadorModel()
        self.view = ConsolidadorView(parent_frame)
        self._vincular_eventos()

    def get_view(self):
        return self.view

    def _vincular_eventos(self):
        # Aba Vendas
        self.view.btn_vendas_proc.configure(command=self.procurar_vendas_path)
        self.view.btn_vendas_gerar.configure(command=self.iniciar_vendas)
        # Aba Mesclar
        self.view.btn_add_files.configure(command=self.adicionar_pdfs_individuais)
        self.view.btn_add_dir.configure(command=self.adicionar_pdfs_pasta)
        self.view.btn_consolidar.configure(command=self.iniciar_mesclagem)
        # Aba Conversor (Word e JPG)
        self.view.btn_proc_word.configure(command=self.procurar_pdf_word)
        self.view.btn_gerar_word.configure(command=self.iniciar_pdf_word)
        self.view.btn_proc_jpg.configure(command=self.procurar_pdf_jpg)
        self.view.btn_gerar_jpg.configure(command=self.iniciar_pdf_jpg)
        # Aba Conversor (JPG para PDF)
        self.view.btn_add_img.configure(command=self.adicionar_imagens)
        self.view.btn_gerar_pdf.configure(command=self.iniciar_jpg_pdf)

    # --- ABA VENDAS ---
    def procurar_vendas_path(self):
        if self.view.modo_leitura_var.get() == "pasta":
            path = filedialog.askdirectory(title="Selecione a pasta de PDFs")
        else:
            path = filedialog.askopenfilename(title="Selecione o arquivo PDF", filetypes=[("Arquivos PDF", "*.pdf")])
        if path:
            self.view.ent_vendas_path.delete(0, "end")
            self.view.ent_vendas_path.insert(0, os.path.normpath(path))

    def iniciar_vendas(self):
        entrada = self.view.ent_vendas_path.get().strip()
        if not entrada:
            messagebox.showerror("Erro", "Selecione uma pasta ou um arquivo PDF de entrada.")
            return
        
        data_doc = self.view.ent_data_doc.get().strip()
        solic = self.view.ent_solicitacao.get().strip()
        self.view.progress.set(0)
        self.view.lbl_porcentagem.configure(text="0%")
        self.view.escrever_log("Iniciando processamento das ordens de venda...")
        
        thread = threading.Thread(target=self.model.processar_vendas, args=(entrada, data_doc, solic, self.view.queue))
        thread.daemon = True
        thread.start()

    # --- ABA MESCLAR ---
    def adicionar_pdfs_individuais(self):
        files = filedialog.askopenfilenames(title="Selecione arquivos PDF", filetypes=[("Arquivos PDF", "*.pdf")])
        if files:
            for f in files:
                f_normalized = os.path.normpath(f)
                if f_normalized not in self.view.arquivos_mescla_lista:
                    self.view.arquivos_mescla_lista.append(f_normalized)
                    self.view.lst_pdfs.insert("end", os.path.basename(f_normalized))
            self.view.escrever_log(f"Adicionados {len(files)} arquivo(s) à lista.")

    def adicionar_pdfs_pasta(self):
        folder = filedialog.askdirectory(title="Selecione a pasta principal")
        if folder:
            encontrados = 0
            for root, dirs, files in os.walk(folder):
                for f in sorted(files):
                    if f.lower().endswith(".pdf") and not f.startswith("~$") and not f.startswith("Mesclado_"):
                        full_path = os.path.normpath(os.path.join(root, f))
                        if full_path not in self.view.arquivos_mescla_lista:
                            self.view.arquivos_mescla_lista.append(full_path)
                            self.view.lst_pdfs.insert("end", f)
                            encontrados += 1
            if encontrados > 0: self.view.escrever_log(f"Adicionados {encontrados} arquivo(s) PDF importados da pasta.")
            else: messagebox.showinfo("Informação", "Nenhum arquivo PDF válido localizado nessa pasta ou subpastas!")

    def iniciar_mesclagem(self):
        if not self.view.arquivos_mescla_lista:
            messagebox.showerror("Erro", "Sua lista de arquivos está vazia!")
            return
        save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("Documento PDF", "*.pdf")], initialfile=f"Mesclado_Ordens_Gree_{dt.now().strftime('%Y%m%d')}.pdf")
        if save_path:
            self.view.progress.set(0)
            self.view.lbl_porcentagem.configure(text="0%")
            self.view.escrever_log("Iniciando processo de mesclagem...")
            thread = threading.Thread(target=self.model.mesclar_pdfs, args=(self.view.arquivos_mescla_lista, os.path.normpath(save_path), self.view.queue))
            thread.daemon = True
            thread.start()

    # --- ABA CONVERSOR ---
    def procurar_pdf_word(self):
        path = filedialog.askopenfilename(title="Selecione o arquivo PDF", filetypes=[("Arquivos PDF", "*.pdf")])
        if path:
            self.view.ent_pdf_word_path.delete(0, "end")
            self.view.ent_pdf_word_path.insert(0, os.path.normpath(path))

    def iniciar_pdf_word(self):
        path = self.view.ent_pdf_word_path.get().strip()
        if not path:
            messagebox.showerror("Erro", "Selecione um arquivo de entrada!")
            return
        self.view.progress.set(0)
        self.view.lbl_porcentagem.configure(text="0%")
        self.view.escrever_log("Iniciando conversão para Word...")
        thread = threading.Thread(target=self.model.converter_pdf_para_word, args=(path, self.view.queue))
        thread.daemon = True
        thread.start()

    def procurar_pdf_jpg(self):
        path = filedialog.askopenfilename(title="Selecione o arquivo PDF", filetypes=[("Arquivos PDF", "*.pdf")])
        if path:
            self.view.ent_pdf_jpg_path.delete(0, "end")
            self.view.ent_pdf_jpg_path.insert(0, os.path.normpath(path))

    def iniciar_pdf_jpg(self):
        path = self.view.ent_pdf_jpg_path.get().strip()
        if not path:
            messagebox.showerror("Erro", "Selecione um arquivo de entrada!")
            return
        dpi_str = self.view.cmb_dpi.get().split()[0]
        try: dpi = int(dpi_str)
        except: dpi = 150
        self.view.progress.set(0)
        self.view.lbl_porcentagem.configure(text="0%")
        self.view.escrever_log("Iniciando extração para JPG...")
        thread = threading.Thread(target=self.model.converter_pdf_para_jpg, args=(path, dpi, self.view.queue))
        thread.daemon = True
        thread.start()

    def adicionar_imagens(self):
        files = filedialog.askopenfilenames(filetypes=[("Imagens", "*.jpg;*.jpeg;*.png")])
        if files:
            for f in files:
                f_normalized = os.path.normpath(f)
                if f_normalized not in self.view.arquivos_jpg_lista:
                    self.view.arquivos_jpg_lista.append(f_normalized)
                    self.view.lst_jpgs.insert("end", os.path.basename(f_normalized))

    def iniciar_jpg_pdf(self):
        if not self.view.arquivos_jpg_lista:
            messagebox.showerror("Erro", "Sua lista está vazia!")
            return
        save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("Documento PDF", "*.pdf")], initialfile=f"Imagens_Compiladas_Gree_{dt.now().strftime('%Y%m%d')}.pdf")
        if save_path:
            self.view.progress.set(0)
            self.view.lbl_porcentagem.configure(text="0%")
            self.view.escrever_log("Iniciando compilação de JPG para PDF...")
            thread = threading.Thread(target=self.model.converter_jpg_para_pdf, args=(self.view.arquivos_jpg_lista, os.path.normpath(save_path), self.view.cmb_orientacao.get(), self.view.cmb_margem.get(), self.view.queue))
            thread.daemon = True
            thread.start()