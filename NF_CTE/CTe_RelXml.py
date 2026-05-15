import os
import xml.etree.ElementTree as ET
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox
from datetime import datetime
import threading

# Configurações de aparência
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class AppCTe(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Extrator de XML CT-e - v1.4")
        self.geometry("600x520")
        self.resizable(False, False)

        self.pasta_origem = ctk.StringVar()
        self.pasta_destino = ctk.StringVar()

        self.setup_ui()

        # Rodapé (Corrigido de 'app' para 'self')
        self.footer_label = ctk.CTkLabel(
            self, 
            text="Desenvolvido por Roberto Santos", 
            font=("Arial", 11), 
            text_color="gray"
        )
        self.footer_label.pack(side="bottom", anchor="w", padx=20, pady=10)

    def setup_ui(self):
        self.label_titulo = ctk.CTkLabel(self, text="Processador de CT-e", font=("Roboto", 24, "bold"))
        self.label_titulo.pack(pady=20)

        # Seleção Origem
        self.frame_origem = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_origem.pack(pady=10, padx=40, fill="x")
        ctk.CTkLabel(self.frame_origem, text="Pasta dos XMLs:", font=("Roboto", 12)).pack(anchor="w")
        self.entry_origem = ctk.CTkEntry(self.frame_origem, textvariable=self.pasta_origem)
        self.entry_origem.pack(side="left", fill="x", expand=True, pady=5)
        self.btn_origem = ctk.CTkButton(self.frame_origem, text="Selecionar", width=100, command=self.selecionar_origem)
        self.btn_origem.pack(side="right", padx=(10, 0), pady=5)

        # Seleção Destino
        self.frame_destino = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_destino.pack(pady=10, padx=40, fill="x")
        ctk.CTkLabel(self.frame_destino, text="Pasta para salvar o Excel:", font=("Roboto", 12)).pack(anchor="w")
        self.entry_destino = ctk.CTkEntry(self.frame_destino, textvariable=self.pasta_destino)
        self.entry_destino.pack(side="left", fill="x", expand=True, pady=5)
        self.btn_destino = ctk.CTkButton(self.frame_destino, text="Selecionar", width=100,
                                         command=self.selecionar_destino)
        self.btn_destino.pack(side="right", padx=(10, 0), pady=5)

        self.progress_bar = ctk.CTkProgressBar(self, width=400)
        self.progress_bar.set(0)
        self.progress_bar.pack(pady=30)

        self.label_status = ctk.CTkLabel(self, text="", font=("Roboto", 11))
        self.label_status.pack()

        self.btn_play = ctk.CTkButton(self, text="▶ INICIAR PROCESSAMENTO", font=("Roboto", 14, "bold"),
                                      height=45, fg_color="#2c3e50", hover_color="#1a252f",
                                      command=self.iniciar_thread)
        self.btn_play.pack(pady=10)

    def selecionar_origem(self):
        caminho = filedialog.askdirectory()
        if caminho: self.pasta_origem.set(caminho)

    def selecionar_destino(self):
        caminho = filedialog.askdirectory()
        if caminho: self.pasta_destino.set(caminho)

    def iniciar_thread(self):
        threading.Thread(target=self.processar_arquivos, daemon=True).start()

    def processar_arquivos(self):
        origem = self.pasta_origem.get()
        destino = self.pasta_destino.get()

        if not origem or not destino:
            messagebox.showwarning("Atenção", "Selecione as pastas de origem e destino!")
            return

        arquivos = [f for f in os.listdir(origem) if f.lower().endswith('.xml')]
        if not arquivos:
            messagebox.showinfo("Informação", "Nenhum arquivo XML encontrado.")
            return

        self.btn_play.configure(state="disabled")
        dados_temporarios = []
        total = len(arquivos)
        ns = {'ns': 'http://www.portalfiscal.inf.br/cte'}

        for idx, arquivo in enumerate(arquivos):
            try:
                tree = ET.parse(os.path.join(origem, arquivo))
                root = tree.getroot()

                def get_tag(path, default=""):
                    element = root.find(path, ns)
                    return element.text if element is not None else default

                # Captura de dados
                n_cte = get_tag('.//ns:nCT')

                raw_date = get_tag('.//ns:dhEmi')
                formatted_date = ""
                if raw_date:
                    date_part = raw_date[:10]
                    formatted_date = datetime.strptime(date_part, '%Y-%m-%d').strftime('%d/%m/%Y')

                nome_emit = get_tag('.//ns:emit/ns:xNome')
                cnpj_emit = get_tag('.//ns:emit/ns:CNPJ')
                nome_rem = get_tag('.//ns:rem/ns:xNome')
                cnpj_rem = get_tag('.//ns:rem/ns:CNPJ')
                cnpj_dest = get_tag('.//ns:dest/ns:CNPJ') or get_tag('.//ns:dest/ns:CPF')
                cidade_dest = get_tag('.//ns:dest/ns:enderDest/ns:xMun')

                # CAPTURA COMO FLOAT (Para garantir que seja exportado como VALOR)
                valor_txt = get_tag('.//ns:vPrest/ns:vTPrest', "0")
                peso_txt = get_tag('.//ns:infCarga/ns:infQ/ns:qCarga', "0")

                v_prest = float(valor_txt.replace(',', '.'))
                peso_bruto = float(peso_txt.replace(',', '.'))

                # Chave (Aspa simples mantida para evitar notação científica no Excel)
                chave_raw = get_tag('.//ns:chCTe')
                if not chave_raw:
                    infCte = root.find(".//ns:infCte", ns)
                    chave_raw = infCte.attrib.get('Id', '')[3:] if infCte is not None else ""
                chave = "'" + chave_raw

                # Cálculos (Permanecem como Float)
                tonelada = peso_bruto / 1000
                v_tonelada = v_prest / tonelada if tonelada > 0 else 0

                dados_temporarios.append({
                    "N_CTE": n_cte,
                    "DATA": formatted_date,
                    "CNPJ_EMIT": cnpj_emit,
                    "EMITENTE": nome_emit,
                    "CNPJ_REMETENTE": cnpj_rem,
                    "NOME_REMETENTE": nome_rem,
                    "DESTINATARIO": cnpj_dest,
                    "CIDADE": cidade_dest,
                    "VALOR": v_prest,  # Salvo como float
                    "PESO_BRUTO": peso_bruto,  # Salvo como float
                    "TONELADA": tonelada,  # Salvo como float
                    "V_TONELADA": v_tonelada,  # Salvo como float
                    "CHAVE_CTE": chave
                })

            except Exception as e:
                print(f"Erro no arquivo {arquivo}: {e}")

            progresso = (idx + 1) / total
            self.progress_bar.set(progresso)
            self.label_status.configure(text=f"Processando {idx + 1} de {total}...")
            self.update_idletasks()

        if dados_temporarios:
            df = pd.DataFrame(dados_temporarios)
            caminho_final = os.path.join(destino, f"Relatorio_CTEs_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")

            try:
                # O Pandas exporta floats como números nativos do Excel
                df.to_excel(caminho_final, index=False)
                messagebox.showinfo("Sucesso", "Mensagem tudo pronto")
            except Exception as e:
                messagebox.showerror("Erro", f"Feche o Excel antes de exportar!\nErro: {e}")

        self.btn_play.configure(state="normal")
        self.progress_bar.set(0)
        self.label_status.configure(text="")


if __name__ == "__main__":
    app = AppCTe()
    app.mainloop()