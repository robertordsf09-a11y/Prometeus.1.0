import pyautogui
import pandas as pd
import openpyxl
import pyperclip
import time
import logging
import os
from pathlib import Path
import customtkinter as ctk
from tkinter import messagebox, filedialog

# =========================================================
# LÓGICA E CONFIGURAÇÕES (MANTIDAS)
# =========================================================
DIRETORIO_SCRIPT = Path(__file__).parent
PASTA_IMAGENS = DIRETORIO_SCRIPT / "Item.Im"
PASTA_ERROS = DIRETORIO_SCRIPT / "Im.Erros"
CAMINHO_LOG = DIRETORIO_SCRIPT / "log_automacao.txt"

PASTA_IMAGENS.mkdir(exist_ok=True)
PASTA_ERROS.mkdir(exist_ok=True)

logging.basicConfig(
    filename=CAMINHO_LOG,
    level=logging.INFO,
    format="%(asctime)s - %(message)s",
    encoding="utf-8",
)


def registrar_log(msg):
    print(msg)
    logging.info(msg)


def formatar_valor_ptbr(valor):
    if pd.isna(valor):
        return ""
    if isinstance(valor, (float, int)):
        if valor == int(valor):
            return str(int(valor))
        return str(valor).replace(".", ",")
    return str(valor)


def verificar_popups_erro():
    arquivos_erro = list(PASTA_ERROS.glob("*.png"))
    for erro_img in arquivos_erro:
        try:
            posicao = pyautogui.locateOnScreen(
                str(erro_img), confidence=0.8, grayscale=True
            )
            if posicao:
                return erro_img.name
        except:
            continue
    return None


def iniciar_automacao(coluna_desejada, nome_imagem, caminho_planilha, janela):
    if not coluna_desejada or not nome_imagem or not caminho_planilha:
        messagebox.showwarning("Atenção", "Preencha todas as etapas antes de começar!")
        return

    caminho_funcao = PASTA_IMAGENS / nome_imagem
    janela.iconify()
    registrar_log(
        f"=== INÍCIO DA SESSÃO | Arquivo: {os.path.basename(caminho_planilha)} ==="
    )
    time.sleep(2)

    try:
        df = pd.read_excel(caminho_planilha, engine="openpyxl")
        itens = df[coluna_desejada].dropna().tolist()
    except Exception as e:
        registrar_log(f"[ERRO] Falha ao ler Excel: {e}")
        janela.deiconify()
        return

    total = len(itens)
    for i, item in enumerate(itens, 1):
        valor = formatar_valor_ptbr(item)
        sucesso = False

        for t in range(1, 6):
            erro_detectado = verificar_popups_erro()
            if erro_detectado:
                registrar_log(f"[STOP] Erro detectado: {erro_detectado}")
                janela.deiconify()
                messagebox.showerror(
                    "Erro no Sistema",
                    f"Automação interrompida!\nDetectado: {erro_detectado}",
                )
                return

            try:
                ponto = pyautogui.locateCenterOnScreen(
                    str(caminho_funcao), confidence=0.8, grayscale=True
                )
                if ponto:
                    pyautogui.click(ponto)
                    time.sleep(1.2)
                    pyperclip.copy(valor)
                    pyautogui.hotkey("ctrl", "v")
                    pyautogui.press("enter")
                    sucesso = True
                    break
            except Exception as e:
                registrar_log(f"Erro na interação: {e}")
            time.sleep(1)

        registrar_log(
            f"[{'SUCESSO' if sucesso else 'FALHA'}] - Item {i}/{total} | Valor: {valor}"
        )

    registrar_log("=== SESSÃO FINALIZADA ===")
    messagebox.showinfo("Sucesso", "O processamento foi concluído!")
    janela.deiconify()


# =========================================================
# INTERFACE GRÁFICA OTIMIZADA (450x500)
# =========================================================


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Janela Compacta
        self.title("Mãos de Mephisto")
        self.geometry("450x500")
        self.resizable(False, False)
        self.caminho_arquivo_selecionado = ""

        ctk.set_appearance_mode("system")
        self.grid_columnconfigure(0, weight=1)

        # --- CABEÇALHO COMPACTO ---
        self.lbl_titulo = ctk.CTkLabel(
            self, text="Mãos de Mephisto", font=("Segoe UI", 20, "bold")
        )
        self.lbl_titulo.grid(row=0, column=0, pady=(20, 2), padx=25, sticky="w")

        self.lbl_desc = ctk.CTkLabel(
            self,
            text="Configure a tarefa abaixo.",
            font=("Segoe UI", 12),
            text_color="gray",
        )
        self.lbl_desc.grid(row=1, column=0, pady=(0, 15), padx=25, sticky="w")

        # --- CARD 1: ARQUIVO ---
        self.card_1 = ctk.CTkFrame(self, corner_radius=10)
        self.card_1.grid(row=2, column=0, padx=20, pady=5, sticky="nsew")

        self.btn_file = ctk.CTkButton(
            self.card_1,
            text="Selecionar Planilha Excel",
            height=32,
            fg_color="#3b3b3b",
            hover_color="#4a4a4a",
            command=self.selecionar_arquivo,
        )
        self.btn_file.pack(fill="x", padx=15, pady=(12, 5))

        self.lbl_status_file = ctk.CTkLabel(
            self.card_1,
            text="Nenhum arquivo carregado",
            font=("Segoe UI", 10),
            text_color="gray",
        )
        self.lbl_status_file.pack(anchor="w", padx=15, pady=(0, 10))

        # --- CARD 2: CONFIGURAÇÃO ---
        self.card_2 = ctk.CTkFrame(self, corner_radius=10)
        self.card_2.grid(row=3, column=0, padx=20, pady=10, sticky="nsew")

        # Dropdowns Otimizados
        imagens = [f.name for f in PASTA_IMAGENS.glob("*.png")]
        self.combo_img = ctk.CTkComboBox(
            self.card_2, values=imagens, width=380, height=32
        )
        self.combo_img.pack(padx=15, pady=(15, 8))
        self.combo_img.set("Selecione a Imagem")

        self.combo_col = ctk.CTkComboBox(
            self.card_2, values=["Aguardando Excel..."], width=380, height=32
        )
        self.combo_col.pack(padx=15, pady=(0, 15))

        # --- BOTÃO PRINCIPAL ---
        self.btn_run = ctk.CTkButton(
            self,
            text="EXECUTAR AGORA",
            height=45,
            corner_radius=22,
            fg_color="#0067c0",
            hover_color="#005aab",
            font=("Segoe UI", 14, "bold"),
            command=self.executar,
        )
        self.btn_run.grid(row=4, column=0, padx=20, pady=(20, 5), sticky="nsew")

        # --- LOG COMPACTO ---
        self.btn_log = ctk.CTkButton(
            self,
            text="Ver Logs",
            width=100,
            fg_color="transparent",
            border_width=1,
            font=("Segoe UI", 11),
            command=lambda: os.startfile(CAMINHO_LOG),
        )
        self.btn_log.grid(row=5, column=0, pady=5)

        # Rodapé
        self.lbl_footer = ctk.CTkLabel(
            self, text="v2.6 • Roberto Santos", text_color="gray", font=("Segoe UI", 9)
        )
        self.lbl_footer.grid(row=6, column=0, pady=(10, 5))

    def selecionar_arquivo(self):
        caminho = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if caminho:
            self.caminho_arquivo_selecionado = caminho
            self.lbl_status_file.configure(
                text=f"✔ {os.path.basename(caminho)}", text_color="#28a745"
            )
            try:
                df_temp = pd.read_excel(caminho, nrows=1)
                colunas = df_temp.columns.tolist()
                self.combo_col.configure(values=colunas)
                if colunas:
                    self.combo_col.set(colunas[0])
            except:
                messagebox.showerror("Erro", "Erro ao ler as colunas.")

    def executar(self):
        img = self.combo_img.get()
        col = self.combo_col.get()
        if (
            "Selecione" in img
            or "Aguardando" in col
            or not self.caminho_arquivo_selecionado
        ):
            messagebox.showwarning(
                "Atenção", "Selecione o arquivo e as opções corretamente."
            )
            return
        iniciar_automacao(col, img, self.caminho_arquivo_selecionado, self)


if __name__ == "__main__":
    app = App()
    app.mainloop()
