from __future__ import annotations

import io
import json
import logging
import os
import queue
import shutil
import sys
import threading
import time
from dataclasses import dataclass, field
from logging.handlers import RotatingFileHandler
from pathlib import Path
from tkinter import filedialog, ttk
from typing import Any

import customtkinter as ctk
import msoffcrypto

# ─── Configurações e Constantes Globais ───────────────────────────────────────

def obter_diretorio_base() -> str:
    """
    Retorna o diretório raiz da aplicação.
    Compatível com execução direta (.py) e executável compilado via Nuitka.
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE_DIR: str = obter_diretorio_base()

# Paleta de Cores (Minimalismo Premium)
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
AZUL_PREMIUM = "#0A84FF"

EXTENSOES_EXCEL = {".xlsx", ".xlsm", ".xls", ".xlsb"}

# ─── Logger Padrão ────────────────────────────────────────────────────────────

def criar_logger(nome_modulo: str, usuario: str = "sistema", caminho_log_especifico: str | None = None) -> logging.Logger:
    """
    Cria logger configurado com formato padrão e rotação de arquivo.
    Salva logs em PROMETEUS_ROOT_DIR/logs/{nome_log}.log.
    """
    usuario_real = os.environ.get("PROMETEUS_USER", usuario)
    dir_base = os.environ.get("PROMETEUS_ROOT_DIR", BASE_DIR)
    nome_log = os.environ.get("PROMETEUS_APP_NAME", "aplicacao")
    
    formato = f"[%(asctime)s],[{usuario_real}],[{nome_modulo}] %(levelname)s: %(message)s"
    formatador = logging.Formatter(formato, datefmt="%Y-%m-%d %H:%M:%S")

    caminho_log = caminho_log_especifico if caminho_log_especifico else os.path.join(dir_base, "logs", f"{nome_log}.log")
    os.makedirs(os.path.dirname(caminho_log), exist_ok=True)

    handler_arquivo = RotatingFileHandler(
        caminho_log, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8"
    )
    handler_arquivo.setFormatter(formatador)

    handler_console = logging.StreamHandler()
    handler_console.setFormatter(formatador)

    logger_inst = logging.getLogger(nome_modulo)
    logger_inst.setLevel(logging.INFO)
    logger_inst.handlers.clear()
    logger_inst.addHandler(handler_arquivo)
    logger_inst.addHandler(handler_console)
    return logger_inst

logger_padrao = criar_logger("excel_protector")

# ─── Modelos e Lógica de Negócio ──────────────────────────────────────────────

@dataclass
class ResultadoProtecao:
    sucesso: bool
    verificado: bool = False
    erro: str = ""
    metodo: str = ""

def preservar_metadados(origem: str, destino: str, logger: logging.Logger) -> None:
    """Copia metadados do arquivo original para o destino."""
    try:
        stat = os.stat(origem)
        os.utime(destino, (stat.st_atime, stat.st_mtime))
        if hasattr(os, 'chmod'):
            os.chmod(destino, stat.st_mode)
    except Exception as e:
        logger.warning("Falha ao preservar metadados para %s: %s", destino, e)

def _verificar_senha(caminho: str, senha: str) -> bool:
    """Tenta abrir o arquivo criptografado com a senha para confirmar proteção."""
    try:
        with open(caminho, "rb") as f:
            office = msoffcrypto.OfficeFile(f)
            office.load_key(password=senha)
            buf = io.BytesIO()
            office.decrypt(buf)
            return buf.tell() > 100
    except Exception:
        return False

def proteger_excel(origem: str, destino: str, senha: str, logger: logging.Logger) -> ResultadoProtecao:
    """Aplica criptografia ECMA-376 ao arquivo Excel e preserva metadados."""
    os.makedirs(os.path.dirname(destino) or ".", exist_ok=True)

    try:
        try:
            with open(origem, "rb") as fin:
                office = msoffcrypto.OfficeFile(fin)
                with open(destino, "wb") as fout:
                    office.encrypt(senha, fout)

            if _verificar_senha(destino, senha):
                preservar_metadados(origem, destino, logger)
                return ResultadoProtecao(True, True, metodo="encrypt_direto")
            
            if os.path.exists(destino):
                os.remove(destino)
        except Exception as e1:
            logger.debug("Tentativa 1 falhou (%s): %s", Path(origem).name, e1)

        buf = io.BytesIO()
        with open(origem, "rb") as fin:
            office = msoffcrypto.OfficeFile(fin)
            office.encrypt(senha, buf)

        dados = buf.getvalue()
        if len(dados) < 100:
            raise RuntimeError("Buffer resultante muito pequeno")

        with open(destino, "wb") as fout:
            fout.write(dados)

        verificado = _verificar_senha(destino, senha)
        preservar_metadados(origem, destino, logger)
        return ResultadoProtecao(True, verificado, metodo="encrypt_buffer")

    except Exception as e2:
        try:
            shutil.copy2(origem, destino)
            return ResultadoProtecao(False, False, erro=str(e2))
        except Exception as e3:
            return ResultadoProtecao(False, False, erro=f"encrypt: {e2} | copy: {e3}")

# ─── Componentes de Interface ─────────────────────────────────────────────────

ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Excel Password Protector v3")
        self.geometry("800x600")
        self.resizable(False, False)
        self.configure(fg_color=FUNDO_PRINCIPAL)

        self._fila_ui: queue.Queue[dict[str, Any]] = queue.Queue()
        
        self.item_checked: dict[str, bool] = {}
        self.item_path: dict[str, str] = {}
        self._processando: bool = False
        self._logger_execucao: logging.Logger = logger_padrao

        self._configurar_grid()
        self._construir_interface()
        self._iniciar_loop_fila()

    def _configurar_grid(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

    def _construir_interface(self) -> None:
        header = ctk.CTkFrame(self, fg_color=SUPERFICIE, height=80, corner_radius=0)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        ctk.CTkLabel(
            header, text="Excel Password Protector", font=ctk.CTkFont(size=22, weight="bold"), text_color=OURO_PRINCIPAL
        ).pack(side="left", padx=30, pady=25)

        main = ctk.CTkFrame(self, fg_color="transparent")
        main.grid(row=1, column=0, sticky="nsew", padx=20, pady=20)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(1, weight=1)

        config = ctk.CTkFrame(main, fg_color=SUPERFICIE, corner_radius=12)
        config.grid(row=0, column=0, sticky="ew", pady=(0, 20))
        
        PADX_PADRAO = 12

        row1 = ctk.CTkFrame(config, fg_color="transparent")
        row1.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(row1, text="Origem:", width=60, anchor="w", text_color=TEXTO_PRIMARIO).pack(side="left")
        self._entry_origem = ctk.CTkEntry(row1, placeholder_text="Pasta de origem...", fg_color=FUNDO_PRINCIPAL, border_color=BORDA_FORTE)
        self._entry_origem.pack(side="left", fill="x", expand=True, padx=PADX_PADRAO)
        ctk.CTkButton(row1, text="Explorar", width=90, fg_color=BORDA_FORTE, hover_color=BORDA_SUTIL, command=self._sel_origem).pack(side="left")
        ctk.CTkButton(row1, text="Escanear", width=90, fg_color=AZUL_PREMIUM, hover_color="#005ecb", command=self._escanear).pack(side="left", padx=(PADX_PADRAO, 0))

        row2 = ctk.CTkFrame(config, fg_color="transparent")
        row2.pack(fill="x", padx=20, pady=5)
        ctk.CTkLabel(row2, text="Destino:", width=60, anchor="w", text_color=TEXTO_PRIMARIO).pack(side="left")
        self._entry_destino = ctk.CTkEntry(row2, placeholder_text="Pasta de destino...", fg_color=FUNDO_PRINCIPAL, border_color=BORDA_FORTE)
        self._entry_destino.pack(side="left", fill="x", expand=True, padx=PADX_PADRAO)
        ctk.CTkButton(row2, text="Explorar", width=90, fg_color=BORDA_FORTE, hover_color=BORDA_SUTIL, command=self._sel_destino).pack(side="left")

        row3 = ctk.CTkFrame(config, fg_color="transparent")
        row3.pack(fill="x", padx=20, pady=(5, 15))
        ctk.CTkLabel(row3, text="Senha:", width=60, anchor="w", text_color=TEXTO_PRIMARIO).pack(side="left")
        self._entry_senha = ctk.CTkEntry(row3, show="●", placeholder_text="Mínimo 4 caracteres", fg_color=FUNDO_PRINCIPAL, border_color=BORDA_FORTE)
        self._entry_senha.pack(side="left", fill="x", expand=True, padx=PADX_PADRAO)
        self._btn_eye = ctk.CTkButton(row3, text="👁", width=40, fg_color=BORDA_FORTE, hover_color=BORDA_SUTIL, command=self._toggle_senha)
        self._btn_eye.pack(side="left")

        tree_container = ctk.CTkFrame(main, fg_color=SUPERFICIE, corner_radius=12)
        tree_container.grid(row=1, column=0, sticky="nsew")
        tree_container.grid_columnconfigure(0, weight=1)
        tree_container.grid_rowconfigure(1, weight=1)

        tree_tools = ctk.CTkFrame(tree_container, fg_color="transparent")
        tree_tools.grid(row=0, column=0, sticky="ew", padx=20, pady=10)
        ctk.CTkLabel(tree_tools, text="Arquivos e Pastas", font=ctk.CTkFont(weight="bold"), text_color=TEXTO_DESTAQUE).pack(side="left")
        ctk.CTkButton(tree_tools, text="Marcar Todos", width=100, height=24, fg_color=BORDA_FORTE, hover_color=BORDA_SUTIL, command=lambda: self._set_all_checks(True)).pack(side="right", padx=5)
        ctk.CTkButton(tree_tools, text="Desmarcar Todos", width=100, height=24, fg_color=BORDA_FORTE, hover_color=BORDA_SUTIL, command=lambda: self._set_all_checks(False)).pack(side="right")

        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", background=FUNDO_PRINCIPAL, foreground=TEXTO_PRIMARIO, fieldbackground=FUNDO_PRINCIPAL, borderwidth=0, font=('Segoe UI', 10), rowheight=25)
        style.map("Treeview", background=[('selected', BORDA_FORTE)])
        
        self.tree = ttk.Treeview(tree_container, columns=("check"), show="tree", selectmode="none")
        self.tree.grid(row=1, column=0, sticky="nsew", padx=(20, 0), pady=(0, 20))
        
        sc = ctk.CTkScrollbar(tree_container, command=self.tree.yview)
        sc.grid(row=1, column=1, sticky="ns", padx=(0, 20), pady=(0, 20))
        self.tree.configure(yscrollcommand=sc.set)

        self.tree.bind("<Button-1>", self._on_tree_click)

        footer = ctk.CTkFrame(self, fg_color=SUPERFICIE, height=120, corner_radius=0)
        footer.grid(row=2, column=0, sticky="ew")
        
        self._bar = ctk.CTkProgressBar(footer, height=8, progress_color=ESMERALDA_PRIMARIA)
        self._bar.pack(fill="x", padx=30, pady=(15, 5))
        self._bar.set(0)

        status_frame = ctk.CTkFrame(footer, fg_color="transparent")
        status_frame.pack(fill="x", padx=30, pady=(0, 10))

        self._lbl_status = ctk.CTkLabel(status_frame, text="Pronto", text_color=TEXTO_SECUNDARIO)
        self._lbl_status.pack(side="left")
        
        self._btn_proc = ctk.CTkButton(status_frame, text="Processar Selecionados", height=36, fg_color=ESMERALDA_PRIMARIA, hover_color=ESMERALDA_DEEP, font=ctk.CTkFont(weight="bold"), command=self._iniciar)
        self._btn_proc.pack(side="right")

        rodape_autoral = ctk.CTkLabel(
            self,
            text="Roberto Santos [LABS]©",
            font=ctk.CTkFont(size=10),
            text_color=TEXTO_SECUNDARIO,
        )
        rodape_autoral.grid(row=3, column=0, sticky="ew", padx=10, pady=(0, 8))

    def _iniciar_loop_fila(self) -> None:
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
        acao = mensagem.get("acao")
        if acao == "atualizar_status":
            texto = mensagem.get("texto", "")
            cor = mensagem.get("cor", TEXTO_SECUNDARIO)
            self._lbl_status.configure(text=texto, text_color=cor)
        elif acao == "atualizar_progresso":
            valor = mensagem.get("valor", 0.0)
            self._bar.set(valor)
        elif acao == "escaneamento_concluido":
            self._lbl_status.configure(text="Escaneamento concluído", text_color=ESMERALDA_SUCESSO)
        elif acao == "processamento_concluido":
            texto = mensagem.get("texto", "")
            self._lbl_status.configure(text=texto, text_color=ESMERALDA_SUCESSO)
            self._btn_proc.configure(state="normal", text="Processar Selecionados")
            self._processando = False

    def _executar_em_thread(self, funcao: Any, *args: Any) -> None:
        threading.Thread(target=funcao, args=args, daemon=True).start()

    def _on_tree_click(self, event: Any) -> None:
        item = self.tree.identify_row(event.y)
        if item:
            new_state = not self.item_checked.get(item, False)
            self._update_item_check(item, new_state)

    def _update_item_check(self, item: str, state: bool) -> None:
        self.item_checked[item] = state
        prefix = "☑" if state else "☐"
        current_text = self.tree.item(item, "text")
        clean_text = current_text[2:] if current_text.startswith(("☑", "☐")) else current_text
        self.tree.item(item, text=f"{prefix} {clean_text}")
        for child in self.tree.get_children(item):
            self._update_item_check(child, state)

    def _set_all_checks(self, state: bool) -> None:
        for item in self.tree.get_children(""):
            self._update_item_check(item, state)

    def _sel_origem(self) -> None:
        p = filedialog.askdirectory()
        if p:
            self._entry_origem.delete(0, "end")
            self._entry_origem.insert(0, p)

    def _sel_destino(self) -> None:
        p = filedialog.askdirectory()
        if p:
            self._entry_destino.delete(0, "end")
            self._entry_destino.insert(0, p)

    def _toggle_senha(self) -> None:
        s = self._entry_senha.cget("show")
        self._entry_senha.configure(show="" if s else "●")
        self._btn_eye.configure(text="🙈" if s else "👁")

    def _escanear(self) -> None:
        origem = self._entry_origem.get().strip()
        if not origem or not os.path.isdir(origem):
            return
        
        self._fila_ui.put({"acao": "atualizar_status", "texto": "Escaneando...", "cor": AZUL_PREMIUM})
        self.tree.delete(*self.tree.get_children())
        self.item_checked.clear()
        self.item_path.clear()
        self._executar_em_thread(self._run_scan, origem)

    def _run_scan(self, raiz_path: str) -> None:
        try:
            nodes: dict[str, str] = {"": ""}
            for root, dirs, files in os.walk(raiz_path):
                rel_root = os.path.relpath(root, raiz_path)
                parent_node = nodes[""] if rel_root == "." else nodes[rel_root]
                
                for d in sorted(dirs):
                    path_d = os.path.join(rel_root, d) if rel_root != "." else d
                    node_id = self.tree.insert(parent_node, "end", text=f"☐ 📁 {d}", open=False)
                    nodes[path_d] = node_id
                    self.item_checked[node_id] = False
                
                for f in sorted(files):
                    is_excel = any(f.lower().endswith(ext) for ext in EXTENSOES_EXCEL)
                    if is_excel:
                        path_f = os.path.join(rel_root, f) if rel_root != "." else f
                        node_id = self.tree.insert(parent_node, "end", text=f"☐ 📄 {f}")
                        self.item_checked[node_id] = False
                        self.item_path[node_id] = path_f
            
            self._fila_ui.put({"acao": "escaneamento_concluido"})
        except Exception:
            logger_padrao.exception("Erro ao escanear diretorio")
            self._fila_ui.put({"acao": "atualizar_status", "texto": "Erro no escaneamento", "cor": ERRO})

    def _iniciar(self) -> None:
        if self._processando:
            return
        senha = self._entry_senha.get().strip()
        origem = self._entry_origem.get().strip()
        destino = self._entry_destino.get().strip()
        selecionados = [path for id_item, path in self.item_path.items() if self.item_checked.get(id_item)]
        
        if not selecionados or len(senha) < 4 or not destino:
            self._fila_ui.put({"acao": "atualizar_status", "texto": "Verifique os campos!", "cor": AVISO})
            return
        
        self._processando = True
        self._btn_proc.configure(state="disabled", text="Processando...")
        self._executar_em_thread(self._processar, origem, destino, senha, selecionados)

    def _processar(self, origem_path: str, destino_path: str, senha: str, selecionados: list[str]) -> None:
        try:
            origem = Path(origem_path)
            ts = time.strftime("%Y%m%d_%H%M%S")
            dest_final = Path(destino_path) / f"Protegido_{origem.name}_{ts}"
            dest_final.mkdir(parents=True, exist_ok=True)
            
            log_path = str(dest_final / f"relatorio_{ts}.log")
            self._logger_execucao = criar_logger("execucao_protecao", caminho_log_especifico=log_path)
            
            todos: list[str] = []
            for r, _, fs in os.walk(origem_path):
                for f in fs:
                    todos.append(os.path.join(r, f))
            
            total = len(todos)
            set_selecionados = set(selecionados)
            
            self._logger_execucao.info("Sessao iniciada — Excel Password Protector v3")
            self._logger_execucao.info("Origem  : %s", origem_path)
            self._logger_execucao.info("Destino : %s", dest_final)
            self._logger_execucao.info("Total arquivos : %d", total)
            self._logger_execucao.info("Marcados protecao: %d", len(set_selecionados))
            
            cnt_prot_ok, cnt_prot_err, cnt_copiados, cnt_erros_cp = 0, 0, 0, 0

            for i, arq_abs in enumerate(todos, 1):
                rel = os.path.relpath(arq_abs, origem_path)
                dest_arq = str(dest_final / rel)
                
                progresso = i / total
                texto_status = f"[{i}/{total}] {os.path.basename(arq_abs)}"
                self._fila_ui.put({"acao": "atualizar_progresso", "valor": progresso})
                self._fila_ui.put({"acao": "atualizar_status", "texto": texto_status, "cor": TEXTO_PRIMARIO})

                if rel in set_selecionados:
                    self._logger_execucao.info("[%d/%d] PROTEGER  %s", i, total, rel)
                    res = proteger_excel(arq_abs, dest_arq, senha, self._logger_execucao)
                    if res.sucesso:
                        cnt_prot_ok += 1
                        self._logger_execucao.info("OK Criptografado (metodo: %s)", res.metodo)
                    else:
                        cnt_prot_err += 1
                        self._logger_execucao.error("ERRO Falha ao proteger: %s", res.erro)
                else:
                    self._logger_execucao.info("[%d/%d] COPIAR    %s", i, total, rel)
                    try:
                        os.makedirs(os.path.dirname(dest_arq), exist_ok=True)
                        shutil.copy2(arq_abs, dest_arq)
                        cnt_copiados += 1
                        self._logger_execucao.info("OK Copiado")
                    except Exception as exc:
                        cnt_erros_cp += 1
                        self._logger_execucao.error("ERRO Falha ao copiar: %s", exc)

            total_erros = cnt_prot_err + cnt_erros_cp
            
            self._logger_execucao.info("RESUMO: Protegidos: %d | Falhas protecao: %d | Copiados: %d | Falhas copia: %d | Erros totais: %d",
                                        cnt_prot_ok, cnt_prot_err, cnt_copiados, cnt_erros_cp, total_erros)
            
            msg = f"Concluído! {cnt_prot_ok} protegidos | {cnt_copiados} copiados | {total_erros} erros"
            self._fila_ui.put({"acao": "processamento_concluido", "texto": msg})
        except Exception:
            logger_padrao.exception("Erro critico no processamento")
            self._fila_ui.put({"acao": "processamento_concluido", "texto": "Erro crítico no processamento"})


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
    logger_padrao.info("Iniciando instância da sub-aplicação Excel Protector")
    try:
        app = App()
        app.mainloop()
        logger_padrao.info("Sub-aplicação Excel Protector encerrada normalmente")
    except Exception:
        logger_padrao.exception("Falha crítica na execução da sub-aplicação Excel Protector")
        sys.exit(1)
