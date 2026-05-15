"""
Excel Password Protector v3 - Final
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
• Proteção por senha com msoffcrypto-tool (criptografia real ECMA-376)
• Preservação de metadados (datas de criação/modificação)
• Visualização em árvore (Treeview) com CHECKBOXES para seleção
• Otimizado para diretórios grandes (milhares de arquivos)
• Log detalhado em arquivo .log na pasta de destino

Dependências:
    pip install customtkinter msoffcrypto-tool Pillow
"""

from __future__ import annotations

import io
import logging
import os
import shutil
import threading
import time
import traceback
from pathlib import Path
from tkinter import filedialog, ttk

import customtkinter as ctk
import msoffcrypto

# ─── Tema ────────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

VERDE    = "#22c55e"
VERMELHO = "#ef4444"
AMARELO  = "#f59e0b"
AZUL     = "#3b82f6"
BG_CARD  = "#1e293b"
BG_MAIN  = "#0f172a"
BG_SCROLL= "#111827"
TEXTO    = "#f1f5f9"
SUBTEXTO = "#94a3b8"

EXTENSOES_EXCEL = {".xlsx", ".xlsm", ".xls", ".xlsb"}

# ─── Logger global ────────────────────────────────────────────────────────────
_logger = logging.getLogger("ExcelProtector")

def configurar_log(caminho_log: str):
    _logger.setLevel(logging.DEBUG)
    _logger.handlers.clear()
    fh = logging.FileHandler(caminho_log, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fmt = logging.Formatter("%(asctime)s [%(levelname)-8s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S")
    fh.setFormatter(fmt)
    _logger.addHandler(fh)
    ch = logging.StreamHandler()
    ch.setLevel(logging.WARNING)
    _logger.addHandler(ch)

# ─── Lógica de proteção e metadados ───────────────────────────────────────────
def preservar_metadados(origem: str, destino: str):
    try:
        stat = os.stat(origem)
        os.utime(destino, (stat.st_atime, stat.st_mtime))
        if hasattr(os, 'chmod'): os.chmod(destino, stat.st_mode)
    except Exception as e:
        _logger.warning("Falha ao preservar metadados para %s: %s", destino, e)

def proteger_excel(origem: str, destino: str, senha: str):
    os.makedirs(os.path.dirname(destino) or ".", exist_ok=True)
    try:
        # Tentativa 1
        with open(origem, "rb") as fin:
            office = msoffcrypto.OfficeFile(fin)
            with open(destino, "wb") as fout:
                office.encrypt(senha, fout)
        
        # Verificação rápida
        with open(destino, "rb") as f:
            off = msoffcrypto.OfficeFile(f)
            off.load_key(password=senha)
            buf = io.BytesIO()
            off.decrypt(buf)
            if buf.tell() > 100:
                preservar_metadados(origem, destino)
                return True
        return False
    except:
        try:
            shutil.copy2(origem, destino)
            return False
        except: return False

# ─── Interface Principal ──────────────────────────────────────────────────────
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Excel Password Protector v3")
        self.geometry("1000x800")
        self.configure(fg_color=BG_MAIN)

        self._pasta_origem = ""
        self._pasta_destino = ""
        self._processando = False
        
        # Mapeamento para controle de seleção: {item_id: boolean}
        self.item_checked = {}
        # Mapeamento para caminhos: {item_id: caminho_relativo}
        self.item_path = {}

        self._setup_ui()

    def _setup_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # Header
        header = ctk.CTkFrame(self, fg_color=BG_CARD, height=80, corner_radius=0)
        header.grid(row=0, column=0, sticky="ew")
        ctk.CTkLabel(header, text="Excel Password Protector", font=ctk.CTkFont(size=22, weight="bold"), text_color=TEXTO).pack(side="left", padx=30)

        # Main
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.grid(row=1, column=0, sticky="nsew", padx=20, pady=20)
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(1, weight=1)

        # Configurações
        config = ctk.CTkFrame(main, fg_color=BG_CARD, corner_radius=12)
        config.grid(row=0, column=0, sticky="ew", pady=(0, 20))
        
        # Origem
        row1 = ctk.CTkFrame(config, fg_color="transparent")
        row1.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(row1, text="Origem:", width=60, anchor="w").pack(side="left")
        self._entry_origem = ctk.CTkEntry(row1, placeholder_text="Pasta de origem...")
        self._entry_origem.pack(side="left", fill="x", expand=True, padx=10)
        ctk.CTkButton(row1, text="Explorar", width=90, command=self._sel_origem).pack(side="left")
        ctk.CTkButton(row1, text="Escanear", width=90, fg_color=AZUL, command=self._escanear).pack(side="left", padx=(10, 0))

        # Destino
        row2 = ctk.CTkFrame(config, fg_color="transparent")
        row2.pack(fill="x", padx=20, pady=5)
        ctk.CTkLabel(row2, text="Destino:", width=60, anchor="w").pack(side="left")
        self._entry_destino = ctk.CTkEntry(row2, placeholder_text="Pasta de destino...")
        self._entry_destino.pack(side="left", fill="x", expand=True, padx=10)
        ctk.CTkButton(row2, text="Explorar", width=90, command=self._sel_destino).pack(side="left")

        # Senha
        row3 = ctk.CTkFrame(config, fg_color="transparent")
        row3.pack(fill="x", padx=20, pady=(5, 15))
        ctk.CTkLabel(row3, text="Senha:", width=60, anchor="w").pack(side="left")
        self._entry_senha = ctk.CTkEntry(row3, show="●", placeholder_text="Mínimo 4 caracteres")
        self._entry_senha.pack(side="left", fill="x", expand=True, padx=10)
        self._btn_eye = ctk.CTkButton(row3, text="👁", width=40, command=self._toggle_senha)
        self._btn_eye.pack(side="left")

        # Treeview Area
        tree_container = ctk.CTkFrame(main, fg_color=BG_CARD, corner_radius=12)
        tree_container.grid(row=1, column=0, sticky="nsew")
        tree_container.grid_columnconfigure(0, weight=1)
        tree_container.grid_rowconfigure(1, weight=1)

        # Toolbar da Árvore
        tree_tools = ctk.CTkFrame(tree_container, fg_color="transparent")
        tree_tools.grid(row=0, column=0, sticky="ew", padx=20, pady=10)
        ctk.CTkLabel(tree_tools, text="Arquivos e Pastas (Selecione para proteger)", font=ctk.CTkFont(weight="bold")).pack(side="left")
        ctk.CTkButton(tree_tools, text="Marcar Todos", width=100, height=24, font=ctk.CTkFont(size=11), command=lambda: self._set_all_checks(True)).pack(side="right", padx=5)
        ctk.CTkButton(tree_tools, text="Desmarcar Todos", width=100, height=24, font=ctk.CTkFont(size=11), command=lambda: self._set_all_checks(False)).pack(side="right")

        # Estilo Treeview
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", background=BG_SCROLL, foreground=TEXTO, fieldbackground=BG_SCROLL, borderwidth=0, font=('Segoe UI', 10), rowheight=25)
        style.map("Treeview", background=[('selected', AZUL)])
        
        self.tree = ttk.Treeview(tree_container, columns=("check"), show="tree", selectmode="none")
        self.tree.grid(row=1, column=0, sticky="nsew", padx=(20, 0), pady=(0, 20))
        
        # Scrollbar
        sc = ctk.CTkScrollbar(tree_container, command=self.tree.yview)
        sc.grid(row=1, column=1, sticky="ns", padx=(0, 20), pady=(0, 20))
        self.tree.configure(yscrollcommand=sc.set)

        # Evento de clique para checkbox
        self.tree.bind("<Button-1>", self._on_tree_click)

        # Footer
        footer = ctk.CTkFrame(self, fg_color=BG_CARD, height=100, corner_radius=0)
        footer.grid(row=2, column=0, sticky="ew")
        
        self._bar = ctk.CTkProgressBar(footer, height=8)
        self._bar.pack(fill="x", padx=30, pady=(15, 5))
        self._bar.set(0)

        self._lbl_status = ctk.CTkLabel(footer, text="Pronto para escanear", text_color=SUBTEXTO)
        self._lbl_status.pack(side="left", padx=30)
        
        self._btn_proc = ctk.CTkButton(footer, text="🚀 Processar Selecionados", height=40, fg_color=VERDE, hover_color="#16a34a", font=ctk.CTkFont(weight="bold"), command=self._iniciar)
        self._btn_proc.pack(side="right", padx=30, pady=15)

    def _on_tree_click(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            # Inverte estado
            new_state = not self.item_checked.get(item, False)
            self._update_item_check(item, new_state)

    def _update_item_check(self, item, state):
        self.item_checked[item] = state
        prefix = "☑" if state else "☐"
        current_text = self.tree.item(item, "text")
        # Remove prefixo antigo se houver
        clean_text = current_text[2:] if current_text.startswith(("☑", "☐")) else current_text
        self.tree.item(item, text=f"{prefix} {clean_text}")
        
        # Propagar para filhos
        for child in self.tree.get_children(item):
            self._update_item_check(child, state)

    def _set_all_checks(self, state):
        for item in self.tree.get_children(""):
            self._update_item_check(item, state)

    def _sel_origem(self):
        p = filedialog.askdirectory()
        if p:
            self._entry_origem.delete(0, "end")
            self._entry_origem.insert(0, p)

    def _sel_destino(self):
        p = filedialog.askdirectory()
        if p:
            self._entry_destino.delete(0, "end")
            self._entry_destino.insert(0, p)

    def _toggle_senha(self):
        s = self._entry_senha.cget("show")
        self._entry_senha.configure(show="" if s else "●")
        self._btn_eye.configure(text="🙈" if s else "👁")

    def _escanear(self):
        origem = self._entry_origem.get().strip()
        if not origem or not os.path.isdir(origem): return
        
        self._lbl_status.configure(text="Escaneando...", text_color=AZUL)
        self.tree.delete(*self.tree.get_children())
        self.item_checked.clear()
        self.item_path.clear()
        
        threading.Thread(target=self._run_scan, args=(origem,), daemon=True).start()

    def _run_scan(self, raiz_path):
        nodes = {"": ""}
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

        self.after(0, lambda: self._lbl_status.configure(text="Escaneamento concluído", text_color=VERDE))

    def _iniciar(self):
        if self._processando: return
        senha = self._entry_senha.get().strip()
        origem = self._entry_origem.get().strip()
        destino = self._entry_destino.get().strip()
        
        selecionados = [path for id, path in self.item_path.items() if self.item_checked.get(id)]
        
        if not selecionados or len(senha) < 4 or not destino:
            self._lbl_status.configure(text="⚠️ Selecione arquivos, destino e senha!", text_color=AMARELO)
            return

        self._processando = True
        self._btn_proc.configure(state="disabled", text="⏳ Processando...")
        threading.Thread(target=self._processar, args=(origem, destino, senha, selecionados), daemon=True).start()

    def _processar(self, origem_path, destino_path, senha, selecionados):
        origem = Path(origem_path)
        ts = time.strftime("%Y%m%d_%H%M%S")
        dest_final = Path(destino_path) / f"Protegido_{origem.name}_{ts}"
        dest_final.mkdir(parents=True, exist_ok=True)
        configurar_log(str(dest_final / "processamento.log"))

        # Lista de todos os arquivos para cópia
        todos = []
        for r, _, fs in os.walk(origem_path):
            for f in fs: todos.append(os.path.join(r, f))
        
        total = len(todos)
        set_selecionados = set(selecionados)
        ok, err, cp = 0, 0, 0

        for i, arq_abs in enumerate(todos, 1):
            rel = os.path.relpath(arq_abs, origem_path)
            dest_arq = str(dest_final / rel)
            self.after(0, lambda v=i/total, t=f"[{i}/{total}] {os.path.basename(arq_abs)}": (self._bar.set(v), self._lbl_status.configure(text=t)))

            if rel in set_selecionados:
                if proteger_excel(arq_abs, dest_arq, senha): ok += 1
                else: err += 1
            else:
                try:
                    os.makedirs(os.path.dirname(dest_arq), exist_ok=True)
                    shutil.copy2(arq_abs, dest_arq)
                    cp += 1
                except: err += 1

        msg = f"✅ Fim! 🔐 {ok} protegidos | 📋 {cp} copiados | ❌ {err} erros"
        self.after(0, self._finalizar, msg)

    def _finalizar(self, msg):
        self._lbl_status.configure(text=msg, text_color=VERDE)
        self._btn_proc.configure(state="normal", text="🚀 Processar Selecionados")
        self._processando = False

if __name__ == "__main__":
    app = App()
    app.mainloop()
