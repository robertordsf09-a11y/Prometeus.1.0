import os
import sys
import time
import queue
import threading
import traceback
import logging
from logging.handlers import RotatingFileHandler
from datetime import datetime
from typing import Any

import pandas as pd
import pyautogui
import customtkinter as ctk
from tkinter import messagebox, filedialog

# --- CONFIGURAÇÃO DE CAMINHOS ---
def obter_diretorio_base() -> str:
    """
    Retorna o diretório raiz da aplicação.
    Compatível com execução direta e executável compilado via Nuitka.
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE_DIR: str = obter_diretorio_base()

def obter_caminho(nome_arquivo: str) -> str:
    """Retorna o caminho absoluto para um arquivo de imagem na pasta base."""
    return os.path.join(BASE_DIR, nome_arquivo)

# --- CORES DO SISTEMA ---
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



# --- LOGS ---
def criar_logger(nome_modulo: str, usuario: str = "sistema") -> logging.Logger:
    """
    Cria logger configurado com formato padrão e rotação de arquivo.
    Salva logs em BASE_DIR/logs/aplicacao.log
    """
    formato = f"[%(asctime)s],[{usuario}],[{nome_modulo}] %(levelname)s: %(message)s"
    formatador = logging.Formatter(formato, datefmt="%Y-%m-%d %H:%M:%S")

    caminho_log = os.path.join(BASE_DIR, "logs", "aplicacao.log")
    os.makedirs(os.path.dirname(caminho_log), exist_ok=True)

    handler_arquivo = RotatingFileHandler(
        caminho_log, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8"
    )
    handler_arquivo.setFormatter(formatador)

    handler_console = logging.StreamHandler()
    handler_console.setFormatter(formatador)

    logger_inst = logging.getLogger(nome_modulo)
    logger_inst.setLevel(logging.INFO)
    
    # Evita duplicação de handlers se a função for chamada novamente
    if not logger_inst.handlers:
        logger_inst.addHandler(handler_arquivo)
        logger_inst.addHandler(handler_console)
        
    return logger_inst

logger = criar_logger("alt_cust")

def registrar_log_antigo(usuario: str, mensagem: str, coluna: str = "", valor: str = "") -> None:
    """Mantém a retrocompatibilidade com o formato de log anterior."""
    data_hora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    log_entry = f"[{data_hora}] Usuário: {usuario} | Coluna: {coluna} | Valor: {valor} | Evento: {mensagem}\n"
    caminho_log = os.path.join(BASE_DIR, "logs", "log_automacao.txt")
    os.makedirs(os.path.dirname(caminho_log), exist_ok=True)
    try:
        with open(caminho_log, "a", encoding="utf-8") as f:
            f.write(log_entry)
    except Exception:
        logger.exception("Falha ao registrar log no arquivo antigo.")

# --- UTILITÁRIOS ---
def formatar_numero_limpo(valor: Any) -> str:
    """Formata valores para o padrão decimal brasileiro (vírgula)."""
    try:
        if pd.isna(valor) or str(valor).strip() == "":
            return ""
        v_str = str(valor).replace(",", ".").strip()
        num = float(v_str)
        return f"{num:.2f}".replace(".", ",")
    except Exception:
        return str(valor).strip()

def esperar_imagem(nome_imagem: str, tempo_limite: int = 60) -> tuple[int, int] | None:
    """Aguarda uma imagem aparecer na tela para prosseguir o loop."""
    caminho = obter_caminho(nome_imagem)
    if not os.path.exists(caminho):
        logger.error("esperar_imagem | imagem não encontrada | caminho=%s", caminho)
        raise FileNotFoundError(f"Imagem não encontrada: {nome_imagem}")
    
    inicio = time.time()
    while (time.time() - inicio) < tempo_limite:
        try:
            posicao = pyautogui.locateCenterOnScreen(caminho, confidence=0.8)
            if posicao:
                return posicao
        except Exception:
            pass
        time.sleep(0.5)
    return None

# --- INTERFACE GRÁFICA (UI) ---
ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")

class AutomacaoApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Automação ERP - Alteração de Custo")
        self.geometry("450x600")
        self.resizable(False, False)
        self.configure(fg_color=FUNDO_PRINCIPAL)

        self.usuario_logado: str = sys.argv[1] if len(sys.argv) > 1 else "SISTEMA"
        self.df_banco: pd.DataFrame | None = None
        self.caminho_arquivo: str = ""
        self.coluna_sel: str = ""

        global logger
        logger = criar_logger("alt_cust", usuario=self.usuario_logado)
        logger.info("Aplicação iniciada. Usuário: %s", self.usuario_logado)

        # Controle de Thread e UI
        self._fila_ui: queue.Queue[dict[str, Any]] = queue.Queue()
        self._evento_resposta = threading.Event()
        self._resposta_usuario: bool = False
        self._em_execucao: bool = False

        self._configurar_grid()
        self._construir_interface()
        self._iniciar_loop_fila()

    def _configurar_grid(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0)  # Rodapé

    def _construir_interface(self) -> None:
        self.frame_principal = ctk.CTkFrame(self, fg_color=SUPERFICIE, corner_radius=14)
        self.frame_principal.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
        self.frame_principal.grid_columnconfigure(0, weight=1)

        self._tela_pre_selecao()

        rodape = ctk.CTkLabel(
            self,
            text="Roberto Santos [LABS]©",
            font=ctk.CTkFont(size=10),
            text_color=TEXTO_SECUNDARIO,
        )
        rodape.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 8))

    def _iniciar_loop_fila(self) -> None:
        """Processa mensagens da fila de UI de forma assíncrona."""
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
        tipo = mensagem.get("tipo")
        if tipo == "pergunta_continuar":
            item = mensagem.get("item", "Desconhecido")
            self.deiconify()
            self.attributes("-topmost", True)
            self.attributes("-topmost", False)
            resposta = messagebox.askyesno(
                "Automação Pausada", 
                f"Item {item} finalizado.\nDeseja continuar para o próximo?\n\n(Clique em 'Não' para fechar o programa)"
            )
            self._resposta_usuario = resposta
            self.iconify()
            self._evento_resposta.set()
            
        elif tipo == "sucesso":
            self.deiconify()
            messagebox.showinfo("Sucesso", "Processamento concluído com sucesso!")
            self._em_execucao = False
            self._tela_pre_selecao()
            
        elif tipo == "erro":
            erro = mensagem.get("conteudo", "Erro desconhecido")
            self.deiconify()
            messagebox.showerror("Erro Crítico", f"Falha na execução:\n\n{erro}")
            self._em_execucao = False
            self._tela_pre_selecao()

    def _limpar_frame(self) -> None:
        for widget in self.frame_principal.winfo_children():
            widget.destroy()



    def _tela_pre_selecao(self) -> None:
        self._limpar_frame()
        
        ctk.CTkLabel(
            self.frame_principal, 
            text=f"👋 Bem-vindo(a),\n{self.usuario_logado}!", 
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color=TEXTO_DESTAQUE
        ).pack(pady=(50, 40))
        
        btn_procurar = ctk.CTkButton(
            self.frame_principal, 
            text="📁 Procurar Arquivo Excel", 
            command=self._abrir_seletor, 
            width=320, 
            height=55, 
            fg_color=OURO_PRINCIPAL,
            hover_color=OURO_ESCURO,
            text_color=FUNDO_PRINCIPAL,
            font=ctk.CTkFont(weight="bold", size=15)
        )
        btn_procurar.pack(padx=36, pady=20)

    def _abrir_seletor(self) -> None:
        caminho = filedialog.askopenfilename(
            filetypes=(("Arquivos Excel", "*.xlsx *.xls *.xlsm"), ("Todos", "*.*"))
        )
        if caminho:
            self.caminho_arquivo = caminho
            try:
                self.df_banco = pd.read_excel(self.caminho_arquivo)
                logger.info("Arquivo Excel carregado: %s", caminho)
                self._tela_seletor_colunas()
            except Exception as e:
                logger.exception("abrir_seletor | falha ao ler excel | caminho=%s", caminho)
                messagebox.showerror("Erro de Leitura", f"Erro ao abrir Excel: {e}")

    def _tela_seletor_colunas(self) -> None:
        if self.df_banco is None:
            return
            
        self._limpar_frame()

        ctk.CTkLabel(
            self.frame_principal, 
            text="✅ Arquivo carregado!", 
            font=ctk.CTkFont(size=18, weight="bold"), 
            text_color=ESMERALDA_SUCESSO
        ).pack(pady=(40, 20))
        
        ctk.CTkLabel(
            self.frame_principal, 
            text="Escolha a coluna de CUSTO:",
            font=ctk.CTkFont(size=14),
            text_color=TEXTO_PRIMARIO
        ).pack(pady=(10, 10))
        
        colunas = list(self.df_banco.columns)
        self.combo_colunas = ctk.CTkComboBox(
            self.frame_principal, 
            values=colunas, 
            state="readonly", 
            width=320,
            height=45,
            fg_color=FUNDO_PRINCIPAL,
            border_color=BORDA_FORTE
        )
        if colunas:
            self.combo_colunas.set("Selecione a coluna...")
        self.combo_colunas.pack(pady=(6, 30), padx=36)

        btn_iniciar = ctk.CTkButton(
            self.frame_principal, 
            text="▶ Iniciar Automação", 
            command=self._iniciar_thread_automacao, 
            width=320, 
            height=55, 
            fg_color=ESMERALDA_PRIMARIA,
            hover_color=ESMERALDA_DEEP,
            font=ctk.CTkFont(weight="bold", size=15)
        )
        btn_iniciar.pack(padx=36)
        
        btn_voltar = ctk.CTkButton(
            self.frame_principal,
            text="Voltar",
            command=self._tela_pre_selecao,
            width=320,
            height=40,
            fg_color="transparent",
            hover_color=BORDA_FORTE,
            text_color=TEXTO_SECUNDARIO
        )
        btn_voltar.pack(pady=15)

    def _iniciar_thread_automacao(self) -> None:
        if self._em_execucao:
            return
            
        self.coluna_sel = self.combo_colunas.get()
        if not self.coluna_sel or self.coluna_sel == "Selecione a coluna...":
            messagebox.showwarning("Aviso", "Selecione a coluna de custo!")
            return

        self._em_execucao = True
        self.iconify()
        threading.Thread(target=self._executar_automacao_worker, daemon=True).start()

    def _executar_automacao_worker(self) -> None:
        if self.df_banco is None:
            self._fila_ui.put({"tipo": "erro", "conteudo": "Dados não carregados."})
            return

        try:
            df_processar = self.df_banco.copy()
            col_item_nome = df_processar.columns[0]
            
            logger.info("Iniciando automação. Coluna alvo: %s", self.coluna_sel)
            registrar_log_antigo(self.usuario_logado, f"Início — Coluna: {self.coluna_sel}")
            time.sleep(2)

            for _, row in df_processar.iterrows():
                item_codigo = str(row[col_item_nome])
                
                logger.info("Processando item: %s", item_codigo)

                # 1. Grupo3
                pos = esperar_imagem("Grupo3.png")
                if not pos: 
                    logger.warning("Imagem Grupo3.png não encontrada. Abortando loop.")
                    break
                pyautogui.click(pos)
                time.sleep(1.5)
                pyautogui.write("3")
                pyautogui.press("enter")
                time.sleep(2)

                # 2. Escrever Item
                pyautogui.write(item_codigo)
                pyautogui.press("enter")
                time.sleep(2)

                # 3. Alteracao
                pos = esperar_imagem("Alteracao.png")
                if not pos: 
                    logger.warning("Imagem Alteracao.png não encontrada. Abortando loop.")
                    break
                pyautogui.click(pos)
                time.sleep(2)

                # O bloco de edição comentado (cue.png e cmup.png)
                # foi mantido removido como no original por segurança,
                # mas o fluxo é o mesmo.

                # Pausa e Verificação de Continuidade
                registrar_log_antigo(self.usuario_logado, "Item OK", self.coluna_sel, item_codigo)
                logger.info("Item processado: %s", item_codigo)
                
                # Pausar e perguntar à UI
                self._evento_resposta.clear()
                self._fila_ui.put({"tipo": "pergunta_continuar", "item": item_codigo})
                self._evento_resposta.wait()  # Aguarda resposta da UI
                
                if not self._resposta_usuario:
                    logger.info("Usuário abortou automação no item: %s", item_codigo)
                    registrar_log_antigo(self.usuario_logado, "Encerrado pelo usuário no popup")
                    self.after(0, self.destroy)
                    return

            self._fila_ui.put({"tipo": "sucesso"})

        except Exception:
            logger.exception("Erro crítico durante a automação.")
            erro_detalhado = traceback.format_exc()
            self._fila_ui.put({"tipo": "erro", "conteudo": erro_detalhado})

if __name__ == "__main__":
    app = AutomacaoApp()
    app.mainloop()