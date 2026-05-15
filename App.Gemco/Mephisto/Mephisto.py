from __future__ import annotations

import os
import sys
import time
import queue
import logging
import threading
from typing import Any, Optional
from logging.handlers import RotatingFileHandler

import pandas as pd
import pyautogui
import pyperclip
import customtkinter as ctk
from tkinter import filedialog, messagebox


# =========================================================
# CONFIGURAÇÕES E CAMINHOS
# =========================================================
def obter_diretorio_base() -> str:
    """
    Retorna o diretório raiz da aplicação.
    Compatível com execução direta (.py) e executável compilado via Nuitka.
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


BASE_DIR: str = obter_diretorio_base()
PASTA_IMAGENS: str = os.path.join(BASE_DIR, "Item.Im")
PASTA_ERROS: str = os.path.join(BASE_DIR, "Im.Erros")

os.makedirs(PASTA_IMAGENS, exist_ok=True)
os.makedirs(PASTA_ERROS, exist_ok=True)


# =========================================================
# LOGGING
# =========================================================
def criar_logger(nome_modulo: str, usuario: str = "sistema") -> logging.Logger:
    """
    Cria logger configurado com formato padrão e rotação de arquivo.
    Salva logs em BASE_DIR/logs/aplicacao.log com rotação a cada 5 MB,
    mantendo até 3 arquivos históricos.
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

    logger = logging.getLogger(nome_modulo)
    logger.setLevel(logging.INFO)
    
    if not logger.hasHandlers():
        logger.addHandler(handler_arquivo)
        logger.addHandler(handler_console)
        
    return logger

logger = criar_logger("mephisto", "roberto")


# =========================================================
# DESIGN TOKENS
# =========================================================
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


# =========================================================
# LÓGICA DE NEGÓCIO (MEPHISTO)
# =========================================================
def formatar_valor_ptbr(valor: Any) -> str:
    """Formata valores numéricos e textos lidos do Excel para inserção."""
    if pd.isna(valor):
        return ""
    if isinstance(valor, (float, int)):
        if valor == int(valor):
            return str(int(valor))
        return str(valor).replace(".", ",")
    return str(valor)


def verificar_popups_erro() -> Optional[str]:
    """Verifica se algum popup de erro catalogado na pasta de erros apareceu na tela."""
    if not os.path.exists(PASTA_ERROS):
        return None
        
    arquivos_erro = [f for f in os.listdir(PASTA_ERROS) if f.lower().endswith(".png")]
    for erro_img in arquivos_erro:
        caminho_completo = os.path.join(PASTA_ERROS, erro_img)
        try:
            posicao = pyautogui.locateOnScreen(
                caminho_completo, confidence=0.8, grayscale=True
            )
            if posicao:
                return erro_img
        except Exception:
            continue
    return None


class ProcessadorMephisto:
    """Responsável por executar a automação de inserção de dados via interface."""
    def __init__(self, caminho_planilha: str, coluna_desejada: str, nome_imagem: str, callback_ui: Any) -> None:
        self._caminho_planilha = caminho_planilha
        self._coluna_desejada = coluna_desejada
        self._nome_imagem = nome_imagem
        self._callback = callback_ui
        self._caminho_funcao = os.path.join(PASTA_IMAGENS, self._nome_imagem)

    def executar(self) -> None:
        """Executa a lógica principal da automação rodando em background."""
        logger.info("=== INÍCIO DA SESSÃO | Arquivo: %s ===", os.path.basename(self._caminho_planilha))
        self._callback({"tipo": "status", "mensagem": "Lendo arquivo Excel..."})
        
        try:
            df = pd.read_excel(self._caminho_planilha, engine="openpyxl")
            itens = df[self._coluna_desejada].dropna().tolist()
        except Exception as erro:
            logger.exception("Falha ao ler Excel: %s", self._caminho_planilha)
            self._callback({"tipo": "erro", "mensagem": f"Falha ao ler Excel: {erro}"})
            return

        total = len(itens)
        if total == 0:
            self._callback({"tipo": "aviso", "mensagem": "Nenhum dado encontrado na coluna selecionada."})
            return

        for i, item in enumerate(itens, 1):
            valor = formatar_valor_ptbr(item)
            sucesso = False
            self._callback({"tipo": "progresso", "atual": i, "total": total, "valor": valor})

            for _ in range(5):
                erro_detectado = verificar_popups_erro()
                if erro_detectado:
                    logger.error("Erro popup detectado na tela: %s", erro_detectado)
                    self._callback({
                        "tipo": "erro_critico", 
                        "mensagem": f"Automação interrompida!\nDetectado popup de erro: {erro_detectado}"
                    })
                    return

                try:
                    ponto = pyautogui.locateCenterOnScreen(
                        self._caminho_funcao, confidence=0.8, grayscale=True
                    )
                    if ponto:
                        pyautogui.click(ponto)
                        time.sleep(1.2)
                        pyperclip.copy(valor)
                        pyautogui.hotkey("ctrl", "v")
                        time.sleep(0.5)
                        pyautogui.press("enter")
                        sucesso = True
                        break
                except Exception as e:
                    logger.warning("Tentativa de clique falhou para o valor '%s': %s", valor, e)
                
                time.sleep(1)

            if sucesso:
                logger.info("SUCESSO - Item %d/%d | Valor: %s", i, total, valor)
            else:
                logger.error("FALHA - Item %d/%d | Valor: %s (Campo não encontrado na tela)", i, total, valor)

        logger.info("=== SESSÃO FINALIZADA ===")
        self._callback({"tipo": "sucesso", "mensagem": "O processamento foi concluído com sucesso!"})


# =========================================================
# INTERFACE GRÁFICA (UI)
# =========================================================
ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")

class AplicacaoPrincipal(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        
        self.title("Mãos de Mephisto")
        self.geometry("450x600")
        self.resizable(False, False)
        self.configure(fg_color=FUNDO_PRINCIPAL)
        
        self._fila_ui: queue.Queue[dict[str, Any]] = queue.Queue()
        self._caminho_arquivo_selecionado: str = ""
        self._processando: bool = False
        
        self._configurar_grid()
        self._construir_interface()
        self._carregar_imagens_iniciais()
        self._iniciar_loop_fila()

    def _configurar_grid(self) -> None:
        """Configura pesos do grid principal da janela."""
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

    def _construir_interface(self) -> None:
        """Constrói todos os componentes visuais de forma modular."""
        # --- CABEÇALHO ---
        frame_header = ctk.CTkFrame(self, fg_color="transparent")
        frame_header.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")
        
        lbl_titulo = ctk.CTkLabel(
            frame_header, 
            text="Mãos de Mephisto", 
            font=ctk.CTkFont(size=24, weight="bold"),
            text_color=OURO_PRINCIPAL
        )
        lbl_titulo.pack(anchor="w")
        
        lbl_desc = ctk.CTkLabel(
            frame_header,
            text="Automação de Entrada de Dados",
            font=ctk.CTkFont(size=14),
            text_color=TEXTO_SECUNDARIO
        )
        lbl_desc.pack(anchor="w")

        # --- CONTEÚDO PRINCIPAL ---
        frame_main = ctk.CTkFrame(self, fg_color=SUPERFICIE, corner_radius=15)
        frame_main.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")
        
        # Seleção de Arquivo
        lbl_arquivo = ctk.CTkLabel(
            frame_main, text="Planilha Base (Excel):", 
            font=ctk.CTkFont(size=12, weight="bold"), text_color=TEXTO_PRIMARIO
        )
        lbl_arquivo.pack(anchor="w", padx=20, pady=(20, 5))
        
        self._btn_arquivo = ctk.CTkButton(
            frame_main,
            text="Selecionar Arquivo",
            fg_color=BORDA_FORTE,
            hover_color=BORDA_SUTIL,
            text_color=TEXTO_DESTAQUE,
            height=40,
            command=self._selecionar_arquivo
        )
        self._btn_arquivo.pack(fill="x", padx=20, pady=5)
        
        self._lbl_status_arquivo = ctk.CTkLabel(
            frame_main, text="Nenhum arquivo carregado", 
            font=ctk.CTkFont(size=12), text_color=TEXTO_SECUNDARIO
        )
        self._lbl_status_arquivo.pack(anchor="w", padx=20, pady=(0, 15))

        # Seleção de Imagem
        lbl_img = ctk.CTkLabel(
            frame_main, text="Imagem do Campo Alvo:", 
            font=ctk.CTkFont(size=12, weight="bold"), text_color=TEXTO_PRIMARIO
        )
        lbl_img.pack(anchor="w", padx=20, pady=(10, 5))
        
        self._combo_img = ctk.CTkComboBox(
            frame_main, 
            values=["Nenhuma imagem encontrada"],
            height=40,
            fg_color=FUNDO_PRINCIPAL,
            border_color=BORDA_FORTE,
            button_color=BORDA_FORTE,
            button_hover_color=BORDA_SUTIL,
            text_color=TEXTO_PRIMARIO
        )
        self._combo_img.pack(fill="x", padx=20, pady=5)

        # Seleção de Coluna
        lbl_col = ctk.CTkLabel(
            frame_main, text="Coluna de Dados (Excel):", 
            font=ctk.CTkFont(size=12, weight="bold"), text_color=TEXTO_PRIMARIO
        )
        lbl_col.pack(anchor="w", padx=20, pady=(15, 5))
        
        self._combo_col = ctk.CTkComboBox(
            frame_main, 
            values=["Aguardando Excel..."],
            height=40,
            fg_color=FUNDO_PRINCIPAL,
            border_color=BORDA_FORTE,
            button_color=BORDA_FORTE,
            button_hover_color=BORDA_SUTIL,
            text_color=TEXTO_PRIMARIO
        )
        self._combo_col.pack(fill="x", padx=20, pady=5)

        # --- ÁREA DE PROGRESSO E CONTROLE ---
        frame_controle = ctk.CTkFrame(self, fg_color="transparent")
        frame_controle.grid(row=2, column=0, padx=20, pady=(10, 20), sticky="ew")

        self._barra_progresso = ctk.CTkProgressBar(
            frame_controle, fg_color=BORDA_FORTE, progress_color=ESMERALDA_PRIMARIA
        )
        self._barra_progresso.pack(fill="x", pady=(0, 15))
        self._barra_progresso.set(0)
        self._barra_progresso.pack_forget()

        self._lbl_status = ctk.CTkLabel(
            frame_controle, text="", 
            font=ctk.CTkFont(size=12, weight="bold"), text_color=ESMERALDA_SUCESSO
        )
        self._lbl_status.pack(pady=(0, 10))

        self._btn_executar = ctk.CTkButton(
            frame_controle,
            text="INICIAR AUTOMAÇÃO",
            height=50,
            corner_radius=25,
            fg_color=ESMERALDA_PRIMARIA,
            hover_color=ESMERALDA_DEEP,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self._iniciar_execucao
        )
        self._btn_executar.pack(fill="x")

        # --- RODAPÉ ---
        rodape = ctk.CTkLabel(
            self,
            text="Roberto Santos [LABS]©",
            font=ctk.CTkFont(size=10),
            text_color=TEXTO_SECUNDARIO,
        )
        rodape.grid(
            row=3,
            column=0,
            sticky="ew",
            padx=10,
            pady=(0, 8),
        )

    def _carregar_imagens_iniciais(self) -> None:
        """Carrega a lista de imagens de referência na interface."""
        if os.path.exists(PASTA_IMAGENS):
            imagens = [f for f in os.listdir(PASTA_IMAGENS) if f.lower().endswith(".png")]
            if imagens:
                self._combo_img.configure(values=imagens)
                self._combo_img.set(imagens[0])
            else:
                self._combo_img.set("Nenhuma imagem encontrada")

    def _iniciar_loop_fila(self) -> None:
        """Inicia a sondagem da fila thread-safe na thread da interface."""
        self._processar_fila()

    def _processar_fila(self) -> None:
        """Processa as mensagens vindas da thread de automação e atualiza a UI."""
        try:
            while True:
                mensagem = self._fila_ui.get_nowait()
                self._tratar_mensagem_ui(mensagem)
        except queue.Empty:
            pass
        self.after(50, self._processar_fila)

    def _tratar_mensagem_ui(self, mensagem: dict[str, Any]) -> None:
        """Aplica os updates de estado ou popups na UI principal."""
        tipo = mensagem.get("tipo")
        
        if tipo == "status":
            self._lbl_status.configure(text=mensagem["mensagem"], text_color=TEXTO_PRIMARIO)
        
        elif tipo == "progresso":
            atual = mensagem["atual"]
            total = mensagem["total"]
            valor = mensagem["valor"]
            self._barra_progresso.set(atual / total if total > 0 else 0)
            self._lbl_status.configure(
                text=f"Processando {atual}/{total} - Valor: {valor}",
                text_color=TEXTO_PRIMARIO
            )
            
        elif tipo == "sucesso":
            self._finalizar_processo()
            self._lbl_status.configure(text=mensagem["mensagem"], text_color=ESMERALDA_SUCESSO)
            messagebox.showinfo("Sucesso", mensagem["mensagem"])
            self._limpar_status(3000)
            
        elif tipo == "erro":
            self._finalizar_processo()
            self._lbl_status.configure(text="Erro no processo.", text_color=ERRO)
            messagebox.showerror("Erro", mensagem["mensagem"])
            
        elif tipo == "erro_critico":
            self._finalizar_processo()
            self._lbl_status.configure(text="Automação interrompida.", text_color=ERRO)
            self.deiconify()
            messagebox.showerror("Erro Crítico", mensagem["mensagem"])
            
        elif tipo == "aviso":
            self._finalizar_processo()
            self._lbl_status.configure(text="Aviso", text_color=AVISO)
            messagebox.showwarning("Atenção", mensagem["mensagem"])

    def _finalizar_processo(self) -> None:
        """Retorna a interface ao seu estado ocioso normal."""
        self._processando = False
        self._barra_progresso.pack_forget()
        self._btn_executar.configure(state="normal", text="INICIAR AUTOMAÇÃO")
        self.deiconify()

    def _limpar_status(self, ms: int) -> None:
        """Limpa o feedback de texto em tela após determinado período."""
        self.after(ms, lambda: self._lbl_status.configure(text=""))

    def _selecionar_arquivo(self) -> None:
        """Processo de seleção via explorador de arquivos e preenchimento das colunas."""
        caminho = filedialog.askopenfilename(
            title="Selecione a Planilha",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
        )
        if not caminho:
            return

        self._caminho_arquivo_selecionado = caminho
        nome_arquivo = os.path.basename(caminho)
        self._lbl_status_arquivo.configure(text=f"✔ {nome_arquivo}", text_color=ESMERALDA_SUCESSO)

        try:
            df_temp = pd.read_excel(caminho, nrows=1, engine="openpyxl")
            colunas = df_temp.columns.tolist()
            if colunas:
                str_colunas = [str(c) for c in colunas]
                self._combo_col.configure(values=str_colunas)
                self._combo_col.set(str_colunas[0])
            else:
                self._combo_col.configure(values=["Nenhuma coluna encontrada"])
                self._combo_col.set("Nenhuma coluna encontrada")
        except Exception as e:
            logger.exception("Falha ao ler colunas do arquivo %s", caminho)
            messagebox.showerror("Erro", f"Não foi possível ler as colunas do Excel.\nDetalhes: {e}")

    def _iniciar_execucao(self) -> None:
        """Validações iniciais e disparador do processamento."""
        if self._processando:
            return

        img = self._combo_img.get()
        col = self._combo_col.get()

        if not self._caminho_arquivo_selecionado:
            messagebox.showwarning("Atenção", "Selecione uma planilha antes de começar.")
            return
        if "Nenhuma" in img or "Selecione" in img:
            messagebox.showwarning("Atenção", "Selecione uma imagem alvo válida.")
            return
        if "Aguardando" in col or "Nenhuma" in col:
            messagebox.showwarning("Atenção", "Selecione a coluna de dados.")
            return

        self._processando = True
        self._btn_executar.configure(state="disabled", text="PROCESSANDO...")
        self._barra_progresso.pack(fill="x", pady=(0, 15))
        self._barra_progresso.set(0)
        
        # Minimizar a janela para a automação rodar livremente sobre o sistema
        self.iconify()
        
        # Aguardar um momento para a janela ser efetivamente ocultada
        time.sleep(1.5)
        
        processador = ProcessadorMephisto(
            caminho_planilha=self._caminho_arquivo_selecionado,
            coluna_desejada=col,
            nome_imagem=img,
            callback_ui=self._fila_ui.put
        )
        
        threading.Thread(target=processador.executar, daemon=True).start()


if __name__ == "__main__":
    try:
        app = AplicacaoPrincipal()
        app.mainloop()
    except Exception as e:
        logger.exception("Falha fatal na aplicação")
        sys.exit(1)
