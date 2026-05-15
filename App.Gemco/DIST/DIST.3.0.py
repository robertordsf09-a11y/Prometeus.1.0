from __future__ import annotations

import logging
import os
import queue
import sys
import threading
import time
import traceback
from dataclasses import dataclass, field
from logging.handlers import RotatingFileHandler
from typing import Any, Callable

import customtkinter as ctk
import pandas as pd
import pyautogui
from tkinter import filedialog, messagebox


# =============================================================================
# CAMINHOS E PORTABILIDADE
# =============================================================================
def obter_diretorio_base() -> str:
    """
    Retorna o diretório raiz da aplicação.

    Compatível com execução direta (.py) e executável compilado via Nuitka.
    Nunca use __file__ diretamente fora desta função.
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


BASE_DIR: str = obter_diretorio_base()


# =============================================================================
# SISTEMA DE LOGS
# =============================================================================
def criar_logger(nome_modulo: str, usuario: str = "sistema") -> logging.Logger:
    """
    Cria logger configurado com formato padrão e rotação de arquivo.
    Salva logs em PROMETEUS_ROOT_DIR/logs/{nome_log}.log.
    """
    usuario_real = os.environ.get("PROMETEUS_USER", usuario)
    dir_base = os.environ.get("PROMETEUS_ROOT_DIR", BASE_DIR)
    nome_log = os.environ.get("PROMETEUS_APP_NAME", "aplicacao")
    
    formato = f"[%(asctime)s],[{usuario_real}],[{nome_modulo}] %(levelname)s: %(message)s"
    formatador = logging.Formatter(formato, datefmt="%Y-%m-%d %H:%M:%S")

    caminho_log = os.path.join(dir_base, "logs", f"{nome_log}.log")
    os.makedirs(os.path.dirname(caminho_log), exist_ok=True)

    handler_arquivo = RotatingFileHandler(
        caminho_log, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8"
    )
    handler_arquivo.setFormatter(formatador)

    handler_console = logging.StreamHandler()
    handler_console.setFormatter(formatador)

    logger_inst = logging.getLogger(nome_modulo)
    logger_inst.setLevel(logging.INFO)
    
    if not logger_inst.handlers:
        logger_inst.addHandler(handler_arquivo)
        logger_inst.addHandler(handler_console)
        
    return logger_inst


logger = criar_logger("dist")


def formatar_numero_limpo(valor: Any) -> str:
    """Formata número para o formato do ERP, removendo decimais zerados e usando vírgula."""
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
# EXCEÇÕES CUSTOMIZADAS
# =============================================================================
class ErroDeAutenticacao(Exception):
    """Levantado quando há falha de login."""

class ErroDeValidacao(Exception):
    """Levantado quando dados são inválidos."""

class ErroDeExecucao(Exception):
    """Levantado quando há erro na automação visual."""


# =============================================================================
# CONSTANTES E CORES DA UI
# =============================================================================
FUNDO_PRINCIPAL = "#0A0A0A"
SUPERFICIE = "#1C1C1C"
BORDA_FORTE = "#2A2A2A"
BORDA_SUTIL = "#3A3A3A"
TEXTO_SECUNDARIO = "#8C8C8C"
TEXTO_PRIMARIO = "#BEBEBE"
TEXTO_DESTAQUE = "#EDEDED"
OURO_PRINCIPAL = "#D4AF37"
OURO_ESCURO = "#B8972E"
ESMERALDA_DEEP = "#006D4E"
ESMERALDA_PRIMARIA = "#00A36C"
ESMERALDA_SUCESSO = "#00C17C"
ERRO = "#C8102E"
PERIGO = "#8B0000"
AVISO = "#FFB800"

PADX_PADRAO = 12
PADY_PADRAO = 10



# =============================================================================
# COMPONENTES DE INTERFACE
# =============================================================================
class BaseUI(ctk.CTk):
    """Base para a janela principal com configurações padrões."""
    def __init__(self) -> None:
        super().__init__()
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")
        
        self.title("Prometeus System - ERP Automation")
        self.geometry("450x600")
        self.resizable(False, False)
        self.configure(fg_color=FUNDO_PRINCIPAL)
        
        self._fila_ui: queue.Queue[dict[str, Any]] = queue.Queue()
        self._iniciar_loop_fila()
        
    def _iniciar_loop_fila(self) -> None:
        """Processa mensagens da fila de UI a cada 50ms."""
        self._processar_fila()

    def _processar_fila(self) -> None:
        try:
            while True:
                mensagem = self._fila_ui.get_nowait()
                self._tratar_mensagem_ui(mensagem)
        except queue.Empty:
            pass
        self.after(50, self._processar_fila)
        
    def _tratar_mensagem_ui(self, mensagem: dict[str, Any]) -> None:
        """Processa as mensagens da fila na thread principal."""
        acao = mensagem.get("acao")
        if acao == "sucesso":
            self.exibir_feedback_sucesso(mensagem.get("mensagem", "Operação concluída!"))
        elif acao == "erro":
            self.exibir_feedback_erro(mensagem.get("mensagem", "Ocorreu um erro."))
        elif acao == "aviso":
            self.exibir_feedback_aviso(mensagem.get("mensagem", "Atenção."))
        elif acao == "callback":
            callback = mensagem.get("callback")
            if callable(callback):
                callback()

    def exibir_feedback_sucesso(self, mensagem: str) -> None:
        """Exibe popup de sucesso."""
        messagebox.showinfo("Sucesso", mensagem)

    def exibir_feedback_erro(self, mensagem: str) -> None:
        """Exibe popup de erro."""
        messagebox.showerror("Erro", mensagem)

    def exibir_feedback_aviso(self, mensagem: str) -> None:
        """Exibe popup de aviso."""
        messagebox.showwarning("Aviso", mensagem)


# =============================================================================
# CLASSE PRINCIPAL
# =============================================================================
class AutomacaoApp(BaseUI):
    """Janela principal da automação."""

    def __init__(self) -> None:
        super().__init__()
        
        self.usuario_logado: str = "SISTEMA"
        self.df_banco: pd.DataFrame | None = None
        self.caminho_arquivo: str = ""
        self.coluna_sel: str = ""

        self._configurar_grid()
        self._construir_interface()

    def _configurar_grid(self) -> None:
        """Configura pesos das colunas e linhas base."""
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0)

    def _construir_interface(self) -> None:
        """Monta o container principal da UI e adiciona o rodapé."""
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(0, weight=1)
        
        rodape = ctk.CTkLabel(
            self,
            text="Roberto Santos [LABS]©",
            font=ctk.CTkFont(size=10),
            text_color=TEXTO_SECUNDARIO,
        )
        rodape.grid(
            row=1,
            column=0,
            sticky="ew",
            padx=PADX_PADRAO,
            pady=(0, 8),
        )

        self.tela_pre_selecao()
    def _limpar_frame(self) -> None:
        """Remove todos os widgets do frame principal."""
        for widget in self.main_frame.winfo_children():
            widget.destroy()

    def _criar_card(self) -> ctk.CTkFrame:
        """Cria um card para colocar os widgets centralizados."""
        card = ctk.CTkFrame(
            self.main_frame, fg_color=SUPERFICIE, corner_radius=16, border_width=1, border_color=BORDA_FORTE
        )
        card.grid(row=0, column=0, padx=30, pady=30, sticky="nsew")
        card.grid_columnconfigure(0, weight=1)
        return card


    def tela_pre_selecao(self) -> None:
        """Monta a tela principal pós-login."""
        self._limpar_frame()
        card = self._criar_card()

        ctk.CTkLabel(
            card,
            text="BEM-VINDO AO NÚCLEO,",
            font=ctk.CTkFont(size=12),
            text_color=TEXTO_SECUNDARIO,
        ).grid(row=0, column=0, pady=(50, 0))

        ctk.CTkLabel(
            card,
            text=self.usuario_logado,
            font=ctk.CTkFont(size=28, weight="bold"),
            text_color=OURO_PRINCIPAL,
        ).grid(row=1, column=0, pady=(0, 50))

        ctk.CTkButton(
            card,
            text="CARREGAR BASE EXCEL",
            command=self.abrir_seletor,
            width=320,
            height=60,
            corner_radius=12,
            fg_color="transparent",
            border_width=2,
            border_color=ESMERALDA_PRIMARIA,
            hover_color=BORDA_FORTE,
            text_color=TEXTO_PRIMARIO,
            font=ctk.CTkFont(size=14, weight="bold"),
        ).grid(row=2, column=0, pady=20)

    def abrir_seletor(self) -> None:
        """Abre caixa de diálogo para carregar arquivo Excel."""
        caminho = filedialog.askopenfilename(
            filetypes=(("Arquivos Excel", "*.xlsx *.xls *.xlsm"), ("Todos", "*.*"))
        )
        if caminho:
            self.caminho_arquivo = caminho
            try:
                self.df_banco = pd.read_excel(self.caminho_arquivo)
                self.tela_seletor_colunas()
            except Exception as erro:
                logger.exception("abrir_seletor | falha na leitura excel | erro=%s", erro)
                self._fila_ui.put({"acao": "erro", "mensagem": f"Erro ao abrir Excel: {erro}"})

    def tela_seletor_colunas(self) -> None:
        """Monta tela para selecionar a coluna da automação."""
        if self.df_banco is None:
            return

        self._limpar_frame()
        card = self._criar_card()

        ctk.CTkLabel(
            card,
            text="DADOS PROCESSADOS",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color=ESMERALDA_PRIMARIA,
        ).grid(row=0, column=0, pady=(40, 5))

        ctk.CTkLabel(
            card,
            text=f"{len(self.df_banco)} registros prontos para envio",
            font=ctk.CTkFont(size=12),
            text_color=TEXTO_SECUNDARIO,
        ).grid(row=1, column=0, pady=(0, 40))

        self.combo_colunas = ctk.CTkComboBox(
            card,
            values=list(self.df_banco.columns),
            width=320,
            height=45,
            corner_radius=12,
            fg_color=FUNDO_PRINCIPAL,
            border_color=OURO_ESCURO,
            button_color=OURO_ESCURO,
            dropdown_hover_color=ESMERALDA_PRIMARIA,
            text_color=TEXTO_DESTAQUE,
        )
        self.combo_colunas.set("Escolha a Coluna")
        self.combo_colunas.grid(row=2, column=0, pady=20)

        self.btn_iniciar = ctk.CTkButton(
            card,
            text="INICIAR AUTOMAÇÃO",
            command=self.iniciar_automacao_thread,
            width=320,
            height=60,
            corner_radius=12,
            fg_color=ESMERALDA_PRIMARIA,
            hover_color=ESMERALDA_DEEP,
            text_color=FUNDO_PRINCIPAL,
            font=ctk.CTkFont(size=14, weight="bold"),
        )
        self.btn_iniciar.grid(row=3, column=0, pady=(20, 30))

        self.barra_progresso = ctk.CTkProgressBar(card, width=320, height=10, progress_color=ESMERALDA_SUCESSO)
        self.barra_progresso.grid(row=4, column=0, pady=10)
        self.barra_progresso.set(0)
        self.barra_progresso.grid_remove()

    def iniciar_automacao_thread(self) -> None:
        """Inicia a automação em uma thread separada para não bloquear a UI."""
        self.coluna_sel = self.combo_colunas.get()
        if not self.coluna_sel or self.coluna_sel == "Escolha a Coluna":
            self._fila_ui.put({"acao": "aviso", "mensagem": "Selecione uma coluna válida."})
            return

        self.btn_iniciar.configure(state="disabled", text="PROCESSANDO...")
        self.barra_progresso.grid()
        self.barra_progresso.configure(mode="indeterminnate")
        self.barra_progresso.start()

        threading.Thread(target=self._executar_automacao_worker, daemon=True).start()

    def _obter_caminho_imagem(self, nome: str) -> str:
        """Retorna caminho para arquivo de imagem (dependência externa)."""
        return os.path.normpath(os.path.join(BASE_DIR, nome))

    def _digitar_texto(self, texto: str) -> None:
        """Função auxiliar de digitação na thread worker."""
        pyautogui.write(str(texto), interval=0.05)
        time.sleep(1.0)
        pyautogui.press("enter")
        time.sleep(2.0)
        
        img_destino = self._obter_caminho_imagem("destino.png")
        if os.path.exists(img_destino):
            for _ in range(20):
                try:
                    if pyautogui.locateOnScreen(img_destino, confidence=0.8):
                        break
                except Exception:
                    pass
                time.sleep(0.5)
    
    '''verificar popup deve verificar se algum dos popups da lista de popups aparece na tela: popup.png, ERRODV.png'''

    def _verificar_popup(self, ultimo_codigo: str) -> bool:
        """Verifica a ocorrência de popup de erro/bloqueio."""
        popups = ["popup.png", "ERRODV.png"]
        for nome_popup in popups:
            img_popup = self._obter_caminho_imagem(nome_popup)
            if os.path.exists(img_popup):
                try:
                    if pyautogui.locateOnScreen(img_popup, confidence=0.6):
                        logger.info(
                            "BLOQUEIO DETECTADO | popup=%s | col=%s | val=%s",
                            nome_popup,
                            self.coluna_sel,
                            ultimo_codigo,
                        )
                        return True
                except Exception:
                    pass
        return False

    def _executar_automacao_worker(self) -> None:
        """Worker thread responsável por manipular pyautogui sem travar a interface."""
        try:
            img_destino = self._obter_caminho_imagem("destino.png")
            img_vazio = self._obter_caminho_imagem("item_vazio.png")

            if not os.path.exists(img_destino) or not os.path.exists(img_vazio):
                msg = (
                    f"Arquivos Críticos Ausentes no diretório:\n{BASE_DIR}\n\n"
                    f"Verifique se 'destino.png' e 'item_vazio.png' estão presentes."
                )
                self._fila_ui.put({"acao": "erro", "mensagem": msg})
                return

            if self.df_banco is None:
                raise ErroDeValidacao("Banco de dados não foi carregado corretamente.")

            df_proc = self.df_banco.copy()
            df_proc[self.coluna_sel] = pd.to_numeric(
                df_proc[self.coluna_sel].astype(str).str.replace(",", "."),
                errors="coerce",
            )
            df_filtrado = df_proc[
                (df_proc[self.coluna_sel].notna()) & (df_proc[self.coluna_sel] != 0)
            ]

            if df_filtrado.empty:
                self._fila_ui.put({"acao": "aviso", "mensagem": "Nenhum valor válido para processar."})
                return

            # Ocultar janela
            self._fila_ui.put({"acao": "callback", "callback": self.iconify})
            
            logger.info(
                "Início da automação | itens=%d | col=%s",
                len(df_filtrado),
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
                    except Exception:
                        pass
                    time.sleep(1.5)

                if pos_vazio:
                    pyautogui.click(pos_vazio)
                    time.sleep(1.0)
                    self._digitar_texto(comando)
                    
                    if self._verificar_popup(comando):
                        self._fila_ui.put({"acao": "erro", "mensagem": f"Interrupção no item: {comando}"})
                        break
                        
                    logger.info(
                        "Item processado com sucesso | col=%s | val=%s",
                        self.coluna_sel,
                        comando,
                    )
                else:
                    self._fila_ui.put({"acao": "erro", "mensagem": f"Campo não limpou para: {comando}"})
                    break

            self._fila_ui.put({"acao": "sucesso", "mensagem": "Fim do processamento!"})

        except Exception as erro:
            logger.exception("_executar_automacao_worker | falha inesperada | erro=%s", erro)
            self._fila_ui.put({"acao": "erro", "mensagem": f"Erro interno: {erro}"})
        finally:
            # Restaurar estado da interface
            def restaurar() -> None:
                self.deiconify()
                self.barra_progresso.stop()
                self.barra_progresso.grid_remove()
                self.btn_iniciar.configure(state="normal", text="INICIAR AUTOMAÇÃO")
                
            self._fila_ui.put({"acao": "callback", "callback": restaurar})


# =============================================================================
# INICIALIZAÇÃO
# =============================================================================
def main() -> None:
    """Ponto de entrada."""
    app = AutomacaoApp()
    app.mainloop()

def validar_execucao_segura() -> None:
    import os
    import sys
    token = os.environ.get("PROMETEUS_AUTH_TOKEN")
    if token != "PR0M3T3U5_L0CK_2026":
        from tkinter import messagebox
        import customtkinter as ctk
        root = ctk.CTk()
        root.withdraw()
        messagebox.showerror(
            "Acesso Negado",
            "Este módulo não pode ser executado isoladamente.\n\n"
            "Por favor, inicie o sistema através do painel principal (Prometeus) e realize o login."
        )
        sys.exit(1)

if __name__ == "__main__":
    validar_execucao_segura()
    logger.info("Iniciando instância da sub-aplicação DIST")
    try:
        main()
        logger.info("Sub-aplicação DIST encerrada normalmente")
    except Exception:
        logger.exception("Falha crítica na execução da sub-aplicação DIST")
        sys.exit(1)

