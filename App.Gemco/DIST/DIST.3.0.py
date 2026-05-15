import sys
import logging
import tkinter as tk
import pandas as pd
import pyautogui
import customtkinter as ctk
from tkinter import messagebox, filedialog
from PIL import Image, ImageTk
import time
from datetime import datetime
import os
import traceback


# =============================================================================
# LICENSE MANAGER (Lógica Original Preservada)
# =============================================================================
class LicenseManager:
    @staticmethod
    def verificar_licenca() -> bool:
        data_expiracao_str = "15/06/2026"
        try:
            data_expiracao = datetime.strptime(data_expiracao_str, "%d/%m/%Y").date()
            if datetime.now().date() <= data_expiracao:
                return True
            LicenseManager._alerta(data_expiracao_str)
            return False
        except Exception as e:
            logging.error(f"Erro na validação de licença: {e}")
            return False

    @staticmethod
    def _alerta(data: str) -> None:
        root = tk.Tk()
        root.withdraw()
        messagebox.showwarning("Licença Expirada", f"Acesso encerrado em {data}.")
        root.destroy()
        sys.exit(0)


# --- CONFIGURAÇÕES DE CAMINHO ---
# Determina o diretório onde o .exe ou o script está localizado
if getattr(sys, "frozen", False) or "__compiled__" in globals():
    # Caminho do executável (.exe)
    BASE_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))
else:
    # Caminho do script (.py)
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

PASTA_IMAGENS = BASE_DIR


def obter_caminho(nome_arquivo: str) -> str:
    """Retorna o caminho absoluto do arquivo dentro da pasta do sistema."""
    return os.path.normpath(os.path.join(PASTA_IMAGENS, nome_arquivo))


# =============================================================================
# PALETA "EMERALD GOTHIC GOLD"
# =============================================================================
# Cores baseadas na preferência "Gótico chic gold/emerald"
BG_DEEP = "#0F0F0F"  # Preto absoluto
BG_FRAME = "#1A1A1A"  # Grafite escuro
EMERALD = "#044D35"  # Verde esmeralda vibrante
EMERALD_HVR = "#06704D"  # Hover esmeralda
GOLD = "#C5A059"  # Ouro metálico
TEXT_MAIN = "#E0E0E0"  # Branco acinzentado
TEXT_MUTED = "#888888"  # Cinza legendas

ctk.set_appearance_mode("dark")


# =============================================================================
# FUNÇÕES DE APOIO (LOG E FORMATAÇÃO)
# =============================================================================
def registrar_log(
    usuario: str, mensagem: str, coluna: str = "", valor: str = ""
) -> None:
    data_hora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    log_entry = (
        f"[{data_hora}] Usuário: {usuario} | Coluna: {coluna} "
        f"| Valor: {valor} | Evento: {mensagem}\n"
    )
    try:
        caminho_log = os.path.join(BASE_DIR, "log_automacao.txt")
        with open(caminho_log, "a", encoding="utf-8") as f:
            f.write(log_entry)
    except Exception as e:
        print(f"Erro ao gravar log: {e}")


def formatar_numero_limpo(valor) -> str:
    try:
        if pd.isna(valor) or str(valor).strip() == "":
            return ""
        v_str = str(valor).replace(",", ".").strip()
        num = float(v_str)
        if num == int(num):
            return str(int(num))
        return str(num).replace(".", ",")
    except Exception:
        return str(valor).strip()


# =============================================================================
# CLASSE PRINCIPAL — PROMETEUS ERP AUTOMATION
# =============================================================================
class AutomacaoApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()

        # Identidade do Projeto
        self.title("Prometeus System - ERP Automation")
        self.geometry("500x600")
        self.resizable(False, False)
        self.configure(fg_color=BG_DEEP)

        # Estado da Aplicação
        self.usuario_logado: str = ""
        self.df_banco = None
        self.caminho_arquivo: str = ""
        self.coluna_sel: str = ""

        self.senhas = {
            "ROBERTO": "rdsf",
            "JOEDSON": "aragao",
            "TAISSA": "fragas",
            "IGOR": "suri",
            "JADSON": "j123",
        }
        self.usuarios_autorizados = list(self.senhas.keys())

        self._configurar_layout_base()
        self.tela_login()

    def _configurar_layout_base(self):
        """Define a estrutura de grid e o rodapé fixo."""
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Rodapé Estilizado com o nome do desenvolvedor
        self.footer = ctk.CTkLabel(
            self,
            text="© 2026 ROBERTO SANTOS | PROMETEUS AUTOMATION",
            font=ctk.CTkFont(family="Segoe UI", size=9, weight="bold"),
            text_color=GOLD,
            fg_color="transparent",
        )
        self.footer.place(relx=0.5, rely=0.96, anchor="center")

    def _criar_card(self):
        """Helper para criar o container central em estilo Card."""
        card = ctk.CTkFrame(
            self, fg_color=BG_FRAME, corner_radius=20, border_width=1, border_color=GOLD
        )
        card.grid(row=0, column=0, padx=30, pady=(40, 60), sticky="nsew")
        card.grid_columnconfigure(0, weight=1)
        return card

    # -------------------------------------------------------------------------
    # TELAS DA INTERFACE
    # -------------------------------------------------------------------------
    def tela_login(self) -> None:
        frame = self._criar_card()

        # Ícone visual (Símbolo Gótico/Ouro)
        ctk.CTkLabel(frame, text="⚜", font=("Segoe UI", 50), text_color=GOLD).grid(
            row=0, column=0, pady=(40, 0)
        )

        ctk.CTkLabel(
            frame,
            text="SISTEMA PROMETEUS",
            font=ctk.CTkFont(
                family="Segoe UI Variable Display", size=22, weight="bold"
            ),
            text_color=TEXT_MAIN,
        ).grid(row=1, column=0, pady=(0, 25))

        self.combo_usuario = ctk.CTkComboBox(
            frame,
            values=self.usuarios_autorizados,
            width=320,
            height=45,
            corner_radius=12,
            fg_color=BG_DEEP,
            border_color=EMERALD,
            button_color=EMERALD,
            text_color=GOLD,
            font=("Segoe UI", 13),
        )
        self.combo_usuario.set("Selecione sua Identidade")
        self.combo_usuario.grid(row=2, column=0, pady=10)

        self.entry_senha = ctk.CTkEntry(
            frame,
            placeholder_text="Senha de Acesso",
            show="●",
            width=320,
            height=45,
            corner_radius=12,
            fg_color=BG_DEEP,
            border_color=EMERALD,
            text_color=TEXT_MAIN,
            placeholder_text_color=TEXT_MUTED,
        )
        self.entry_senha.grid(row=3, column=0, pady=10)
        self.entry_senha.bind("<Return>", lambda _: self.validar_login())

        ctk.CTkButton(
            frame,
            text="AUTENTICAR  →",
            command=self.validar_login,
            width=320,
            height=50,
            corner_radius=12,
            fg_color=EMERALD,
            hover_color=EMERALD_HVR,
            text_color=GOLD,
            font=ctk.CTkFont(family="Segoe UI", weight="bold"),
        ).grid(row=4, column=0, pady=(30, 20))

    def tela_pre_selecao(self) -> None:
        frame = self._criar_card()

        ctk.CTkLabel(
            frame,
            text="BEM-VINDO AO NÚCLEO,",
            font=("Segoe UI", 12),
            text_color=TEXT_MUTED,
        ).grid(row=0, column=0, pady=(50, 0))

        ctk.CTkLabel(
            frame,
            text=self.usuario_logado,
            font=("Segoe UI Variable Display", 28, "bold"),
            text_color=GOLD,
        ).grid(row=1, column=0, pady=(0, 50))

        ctk.CTkButton(
            frame,
            text="📂   CARREGAR BASE EXCEL",
            command=self.abrir_seletor,
            width=340,
            height=65,
            corner_radius=15,
            fg_color="transparent",
            border_width=2,
            border_color=EMERALD,
            hover_color=BG_DEEP,
            text_color=TEXT_MAIN,
            font=ctk.CTkFont(size=14, weight="bold"),
        ).grid(row=2, column=0, pady=20)

    def tela_seletor_colunas(self) -> None:
        frame = self._criar_card()

        ctk.CTkLabel(
            frame,
            text="DADOS PROCESSADOS",
            font=("Segoe UI", 20, "bold"),
            text_color=EMERALD,
        ).grid(row=0, column=0, pady=(40, 5))

        ctk.CTkLabel(
            frame,
            text=f"{len(self.df_banco)} registros prontos para envio",
            font=("Segoe UI", 12),
            text_color=TEXT_MUTED,
        ).grid(row=1, column=0, pady=(0, 40))

        self.combo_colunas = ctk.CTkComboBox(
            frame,
            values=list(self.df_banco.columns),
            width=340,
            height=45,
            corner_radius=12,
            fg_color=BG_DEEP,
            border_color=GOLD,
            button_color=GOLD,
            dropdown_hover_color=EMERALD,
            text_color=TEXT_MAIN,
        )
        self.combo_colunas.set("Escolha a Coluna")
        self.combo_colunas.grid(row=2, column=0, pady=20)

        ctk.CTkButton(
            frame,
            text="⚡   INICIAR AUTOMAÇÃO",
            command=self.executar_automacao,
            width=340,
            height=60,
            corner_radius=12,
            fg_color=EMERALD,
            hover_color=EMERALD_HVR,
            text_color=GOLD,
            font=ctk.CTkFont(size=14, weight="bold"),
        ).grid(row=3, column=0, pady=(20, 30))

    # -------------------------------------------------------------------------
    # LÓGICA OPERACIONAL
    # -------------------------------------------------------------------------
    def validar_login(self) -> None:
        nome = self.combo_usuario.get()
        if nome == "Selecione sua Identidade" or nome not in self.usuarios_autorizados:
            messagebox.showwarning("Aviso", "Selecione um usuário válido.")
            return
        senha = self.entry_senha.get()
        if senha != self.senhas.get(nome):
            messagebox.showwarning("Acesso Negado", "Senha incorreta.")
            self.entry_senha.delete(0, "end")
            return
        self.usuario_logado = nome
        registrar_log(nome, "Login realizado")
        self.tela_pre_selecao()

    def abrir_seletor(self) -> None:
        caminho = filedialog.askopenfilename(
            filetypes=(("Arquivos Excel", "*.xlsx *.xls *.xlsm"), ("Todos", "*.*"))
        )
        if caminho:
            self.caminho_arquivo = caminho
            try:
                self.df_banco = pd.read_excel(self.caminho_arquivo)
                self.tela_seletor_colunas()
            except Exception as e:
                messagebox.showerror("Erro de Leitura", f"Erro ao abrir Excel: {e}")

    def digitar_texto(self, texto: str) -> None:
        pyautogui.write(str(texto), interval=0.05)
        time.sleep(1.0)
        pyautogui.press("enter")
        time.sleep(2.0)
        img_destino = obter_caminho("destino.png")
        if os.path.exists(img_destino):
            for _ in range(20):
                try:
                    if pyautogui.locateOnScreen(img_destino, confidence=0.8):
                        break
                except:
                    pass
                time.sleep(0.5)

    def verificar_popup(self, ultimo_codigo: str) -> bool:
        img_popup = obter_caminho("popup.png")
        if os.path.exists(img_popup):
            try:
                if pyautogui.locateOnScreen(img_popup, confidence=0.6):
                    registrar_log(
                        self.usuario_logado,
                        "BLOQUEIO: Popup",
                        self.coluna_sel,
                        ultimo_codigo,
                    )
                    return True
            except:
                pass
        return False

    def executar_automacao(self) -> None:
        try:
            self.coluna_sel = self.combo_colunas.get()
            if not self.coluna_sel or self.coluna_sel == "Escolha a Coluna":
                messagebox.showwarning("Aviso", "Selecione uma coluna!")
                return

            img_destino = obter_caminho("destino.png")
            img_vazio = obter_caminho("item_vazio.png")

            if not os.path.exists(img_destino) or not os.path.exists(img_vazio):
                msg = (
                    f"Arquivos Críticos Ausentes na pasta:\n{PASTA_IMAGENS}\n\n"
                    f"Verifique se 'destino.png' e 'item_vazio.png' estão presentes."
                )
                messagebox.showerror("Erro de Distribuição", msg)
                return

            df_proc = self.df_banco.copy()
            df_proc[self.coluna_sel] = pd.to_numeric(
                df_proc[self.coluna_sel].astype(str).str.replace(",", "."),
                errors="coerce",
            )
            df_filtrado = df_proc[
                (df_proc[self.coluna_sel].notna()) & (df_proc[self.coluna_sel] != 0)
            ]

            if df_filtrado.empty:
                messagebox.showinfo("Fim", "Nenhum valor válido para processar.")
                return

            self.iconify()
            registrar_log(
                self.usuario_logado,
                f"Início — Itens: {len(df_filtrado)}",
                self.coluna_sel,
            )
            time.sleep(2)

            pos_destino = pyautogui.locateCenterOnScreen(img_destino, confidence=0.8)
            if pos_destino:
                pyautogui.click(pos_destino)
                time.sleep(0.5)
                pyautogui.write(str(self.coluna_sel), interval=0.1)
                pyautogui.press("enter")

            col_item_nome = self.df_banco.columns[0]
            for idx, row in df_filtrado.iterrows():
                qtd = formatar_numero_limpo(row[self.coluna_sel])
                item = formatar_numero_limpo(self.df_banco.loc[idx, col_item_nome])
                comando = f"{qtd}*{item}"

                pos_vazio = None
                inicio_busca = time.time()
                while (time.time() - inicio_busca) < 90:
                    try:
                        pos_vazio = pyautogui.locateCenterOnScreen(
                            img_vazio, confidence=0.9
                        )
                        if pos_vazio:
                            break
                    except:
                        pass
                    time.sleep(1.5)

                if pos_vazio:
                    pyautogui.click(pos_vazio)
                    time.sleep(1.0)
                    self.digitar_texto(comando)
                    if self.verificar_popup(comando):
                        messagebox.showerror(
                            "Erro ERP", f"Interrupção no item: {comando}"
                        )
                        break
                    registrar_log(
                        self.usuario_logado, "Item OK", self.coluna_sel, comando
                    )
                else:
                    messagebox.showerror(
                        "Erro de Sincronia", f"Campo não limpou para: {comando}"
                    )
                    break

            messagebox.showinfo("Sucesso", "Fim do processamento!")
        except Exception:
            messagebox.showerror("Erro Inesperado", traceback.format_exc())
        finally:
            self.deiconify()


# =============================================================================
# INICIALIZAÇÃO
# =============================================================================
if __name__ == "__main__":
    if LicenseManager.verificar_licenca():
        app = AutomacaoApp()
        app.mainloop()
