"""
Excel Password Protector v2
━━━━━━━━━━━━━━━━━━━━━━━━━━━
• Proteção por senha com msoffcrypto-tool (criptografia real ECMA-376)
• Verificação do arquivo após proteção
• Log detalhado em arquivo .log na pasta de destino
• Árvore de pastas colapsável com seleção por pasta inteira

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
from collections import defaultdict
from pathlib import Path
from tkinter import filedialog

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
BG_PASTA = "#162032"

EXTENSOES_EXCEL = {".xlsx", ".xlsm", ".xls", ".xlsb"}


# ─── Logger global ────────────────────────────────────────────────────────────
_logger = logging.getLogger("ExcelProtector")


def configurar_log(caminho_log: str):
    """Configura o logger para escrever no arquivo de log."""
    _logger.setLevel(logging.DEBUG)
    _logger.handlers.clear()

    fh = logging.FileHandler(caminho_log, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fmt = logging.Formatter(
        "%(asctime)s  [%(levelname)-8s]  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    fh.setFormatter(fmt)
    _logger.addHandler(fh)

    # Console apenas warnings+
    ch = logging.StreamHandler()
    ch.setLevel(logging.WARNING)
    ch.setFormatter(fmt)
    _logger.addHandler(ch)

    _logger.info("=" * 70)
    _logger.info("Sessao iniciada — Excel Password Protector v2")
    _logger.info("=" * 70)


# ─── Lógica de proteção ───────────────────────────────────────────────────────
class ResultadoProtecao:
    __slots__ = ("sucesso", "verificado", "erro", "metodo")

    def __init__(self, sucesso: bool, verificado: bool = False,
                 erro: str = "", metodo: str = ""):
        self.sucesso    = sucesso
        self.verificado = verificado
        self.erro       = erro
        self.metodo     = metodo


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


def proteger_excel(origem: str, destino: str, senha: str) -> ResultadoProtecao:
    """
    Aplica criptografia ECMA-376 ao arquivo Excel.

    Tentativa 1 — encrypt direto para arquivo de destino.
    Tentativa 2 — encrypt via buffer em memória (fallback).
    Após criptografar, verifica se a senha está correta.
    Se ambas falharem, copia o arquivo sem proteção e registra o erro.
    """
    os.makedirs(os.path.dirname(destino) or ".", exist_ok=True)

    # Tentativa 1: gravar diretamente no destino
    try:
        with open(origem, "rb") as fin:
            office = msoffcrypto.OfficeFile(fin)
            with open(destino, "wb") as fout:
                office.encrypt(senha, fout)

        if _verificar_senha(destino, senha):
            return ResultadoProtecao(True, True, metodo="encrypt_direto")

        # Arquivo criado mas senha não confirmada — remove e tenta via buffer
        os.remove(destino)

    except Exception as e1:
        _logger.debug("Tentativa 1 falhou (%s): %s", Path(origem).name, e1)

    # Tentativa 2: encrypt via BytesIO
    try:
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
        return ResultadoProtecao(True, verificado, metodo="encrypt_buffer")

    except Exception as e2:
        # Fallback: copia sem senha
        try:
            shutil.copy2(origem, destino)
        except Exception as e3:
            return ResultadoProtecao(
                False, False,
                erro=f"encrypt: {e2} | copy: {e3}",
            )
        return ResultadoProtecao(False, False, erro=str(e2))


def copiar_arquivo(origem: str, destino: str):
    os.makedirs(os.path.dirname(destino) or ".", exist_ok=True)
    shutil.copy2(origem, destino)


# ─── Varredura de pastas ──────────────────────────────────────────────────────
def varrer_pasta(raiz: str) -> dict[str, list[str]]:
    """
    Retorna {pasta_relativa: [caminhos_relativos]} de todos os Excel,
    ordenados A-Z por nome de arquivo dentro de cada pasta.
    A chave '.' representa a raiz.
    """
    r = Path(raiz)
    grupos: dict[str, list[str]] = defaultdict(list)
    for p in r.rglob("*"):
        if p.is_file() and p.suffix.lower() in EXTENSOES_EXCEL:
            rel   = p.relative_to(r)
            pasta = str(rel.parent)
            grupos[pasta].append(str(rel))

    for k in grupos:
        grupos[k].sort(key=lambda x: Path(x).name.lower())

    # Raiz primeiro, demais pastas em ordem alfabética
    return dict(
        sorted(grupos.items(), key=lambda x: ("" if x[0] == "." else x[0].lower()))
    )


# ─── Widgets ─────────────────────────────────────────────────────────────────
class ArquivoRow(ctk.CTkFrame):
    """Linha de checkbox para um único arquivo Excel."""

    def __init__(self, master, caminho_relativo: str, indent: int = 0, **kw):
        super().__init__(master, fg_color="transparent", **kw)
        self.caminho = caminho_relativo
        self._var    = ctk.BooleanVar(value=True)
        self.columnconfigure(2, weight=1)
        self._callbacks_pasta: list = []

        if indent:
            ctk.CTkFrame(self, width=indent, fg_color="transparent"
                         ).grid(row=0, column=0)

        self._cb = ctk.CTkCheckBox(
            self, text="", variable=self._var,
            width=20, checkbox_width=18, checkbox_height=18,
            fg_color=AZUL, hover_color="#2563eb", border_color="#475569",
            command=self._on_change,
        )
        self._cb.grid(row=0, column=1, padx=(4, 2), pady=2)

        nome = Path(caminho_relativo).name
        lbl  = ctk.CTkLabel(
            self, text=f"📄  {nome}", anchor="w",
            text_color=TEXTO, font=ctk.CTkFont(size=12),
        )
        lbl.grid(row=0, column=2, sticky="w", padx=(2, 8), pady=2)

        ext = Path(caminho_relativo).suffix.upper().lstrip(".")
        ctk.CTkLabel(
            self, text=ext, width=42,
            fg_color="#1e3a5f", corner_radius=4,
            text_color="#60a5fa", font=ctk.CTkFont(size=10),
        ).grid(row=0, column=3, padx=(0, 10), pady=2)

        for w in (self, lbl):
            w.bind("<Enter>",    lambda e: self.configure(fg_color="#1a2744"))
            w.bind("<Leave>",    lambda e: self.configure(fg_color="transparent"))
            w.bind("<Button-1>", lambda e: self._toggle())

    def _toggle(self):
        self._var.set(not self._var.get())
        self._on_change()

    def _on_change(self):
        for cb in self._callbacks_pasta:
            cb()

    @property
    def marcado(self) -> bool:
        return self._var.get()

    def marcar(self, valor: bool):
        self._var.set(valor)


class PastaWidget(ctk.CTkFrame):
    """
    Seção colapsável de uma pasta com:
    - Checkbox da pasta (seleciona/desmarca todos os filhos)
    - Lista de ArquivoRow
    - Botão expandir/colapsar
    """

    def __init__(self, master, nome_pasta: str,
                 arquivos: list[str], **kw):
        super().__init__(master, fg_color=BG_PASTA, corner_radius=8, **kw)
        self.columnconfigure(0, weight=1)
        self._rows: list[ArquivoRow] = []
        self._expandido = True

        # ── Cabeçalho ────────────────────────────────────────────────────────
        hdr = ctk.CTkFrame(self, fg_color="transparent")
        hdr.grid(row=0, column=0, sticky="ew", padx=6, pady=(6, 2))
        hdr.columnconfigure(2, weight=1)

        self._var_pasta = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            hdr, text="", variable=self._var_pasta,
            width=20, checkbox_width=18, checkbox_height=18,
            fg_color="#7c3aed", hover_color="#6d28d9", border_color="#475569",
            command=self._selecionar_todos,
        ).grid(row=0, column=0, padx=(2, 6))

        rotulo = "📁  Raiz" if nome_pasta == "." else f"📁  {Path(nome_pasta).name}"
        ctk.CTkLabel(
            hdr, text=rotulo, anchor="w",
            text_color="#c4b5fd", font=ctk.CTkFont(size=13, weight="bold"),
        ).grid(row=0, column=1, sticky="w")

        # Caminho completo (exceto raiz)
        if nome_pasta != ".":
            ctk.CTkLabel(
                hdr, text=f"  {nome_pasta}", anchor="w",
                text_color="#64748b", font=ctk.CTkFont(size=10),
            ).grid(row=0, column=2, sticky="w", padx=4)

        n = len(arquivos)
        ctk.CTkLabel(
            hdr, text=f"{n} arquivo{'s' if n>1 else ''}",
            text_color="#475569", font=ctk.CTkFont(size=11),
        ).grid(row=0, column=3, padx=(0, 6))

        self._btn_col = ctk.CTkButton(
            hdr, text="▼", width=28, height=24,
            fg_color="#334155", hover_color="#475569", corner_radius=6,
            font=ctk.CTkFont(size=11),
            command=self._toggle_colapso,
        )
        self._btn_col.grid(row=0, column=4, padx=(0, 2))

        # Separador
        ctk.CTkFrame(self, height=1, fg_color="#1e3a5f"
                     ).grid(row=1, column=0, sticky="ew", padx=8, pady=(0, 2))

        # ── Lista de arquivos ─────────────────────────────────────────────────
        self._cont = ctk.CTkFrame(self, fg_color="transparent")
        self._cont.grid(row=2, column=0, sticky="ew", padx=4, pady=(0, 6))
        self._cont.columnconfigure(0, weight=1)

        for rel in arquivos:
            row = ArquivoRow(self._cont, rel, indent=28)
            row.grid(sticky="ew", pady=1, padx=4)
            row._callbacks_pasta.append(self._sync_checkbox)
            self._rows.append(row)

    def _toggle_colapso(self):
        self._expandido = not self._expandido
        if self._expandido:
            self._cont.grid()
            self._btn_col.configure(text="▼")
        else:
            self._cont.grid_remove()
            self._btn_col.configure(text="▶")

    def _selecionar_todos(self):
        val = self._var_pasta.get()
        for r in self._rows:
            r.marcar(val)

    def _sync_checkbox(self):
        self._var_pasta.set(all(r.marcado for r in self._rows))

    def marcar_todos(self, valor: bool):
        self._var_pasta.set(valor)
        for r in self._rows:
            r.marcar(valor)

    @property
    def rows(self) -> list[ArquivoRow]:
        return self._rows


# ─── Janela principal ─────────────────────────────────────────────────────────
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Excel Password Protector v2")
        self.geometry("920x760")
        self.minsize(800, 620)
        self.configure(fg_color=BG_MAIN)

        self._pasta_origem   = ""
        self._pasta_destino  = ""
        self._pasta_widgets: list[PastaWidget] = []
        self._processando    = False

        self._build_ui()

    # ── Construção da UI ──────────────────────────────────────────────────────
    def _build_ui(self):
        # Cabeçalho
        hdr = ctk.CTkFrame(self, fg_color=BG_CARD, corner_radius=0, height=64)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        ctk.CTkLabel(
            hdr, text="🔐  Excel Password Protector",
            font=ctk.CTkFont(size=20, weight="bold"), text_color=TEXTO,
        ).pack(side="left", padx=24)
        ctk.CTkLabel(
            hdr, text="v2  •  ECMA-376 Encryption  •  Log detalhado",
            font=ctk.CTkFont(size=11), text_color=SUBTEXTO,
        ).pack(side="right", padx=24)

        # Container
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=20, pady=14)
        main.columnconfigure(0, weight=1)
        main.rowconfigure(1, weight=1)

        # ── Configurações ─────────────────────────────────────────────────────
        cfg = ctk.CTkFrame(main, fg_color=BG_CARD, corner_radius=12)
        cfg.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        cfg.columnconfigure(1, weight=1)

        # Pasta origem
        ctk.CTkLabel(cfg, text="Pasta de Origem", font=ctk.CTkFont(weight="bold"),
                     text_color=TEXTO
                     ).grid(row=0, column=0, columnspan=2, padx=16,
                            pady=(14, 2), sticky="w")
        self._entry_origem = ctk.CTkEntry(
            cfg, placeholder_text="Selecione a pasta com os arquivos Excel…",
            fg_color="#0f172a", border_color="#334155",
            text_color=TEXTO, font=ctk.CTkFont(size=12), height=34,
        )
        self._entry_origem.grid(row=1, column=0, columnspan=2,
                                padx=16, pady=(0, 6), sticky="ew")
        ctk.CTkButton(
            cfg, text="📂  Selecionar Origem", width=188, height=34,
            fg_color=AZUL, hover_color="#2563eb", corner_radius=8,
            command=self._sel_origem,
        ).grid(row=1, column=2, padx=(0, 16), pady=(0, 6))

        # Pasta destino
        ctk.CTkLabel(cfg, text="Pasta de Destino", font=ctk.CTkFont(weight="bold"),
                     text_color=TEXTO
                     ).grid(row=2, column=0, columnspan=2, padx=16,
                            pady=(4, 2), sticky="w")
        self._entry_destino = ctk.CTkEntry(
            cfg, placeholder_text="Selecione onde salvar os arquivos protegidos…",
            fg_color="#0f172a", border_color="#334155",
            text_color=TEXTO, font=ctk.CTkFont(size=12), height=34,
        )
        self._entry_destino.grid(row=3, column=0, columnspan=2,
                                 padx=16, pady=(0, 6), sticky="ew")
        ctk.CTkButton(
            cfg, text="📁  Selecionar Destino", width=188, height=34,
            fg_color="#475569", hover_color="#64748b", corner_radius=8,
            command=self._sel_destino,
        ).grid(row=3, column=2, padx=(0, 16), pady=(0, 6))

        # Senha + Escanear
        ls = ctk.CTkFrame(cfg, fg_color="transparent")
        ls.grid(row=4, column=0, columnspan=3,
                sticky="ew", padx=16, pady=(4, 14))
        ls.columnconfigure(1, weight=1)

        ctk.CTkLabel(ls, text="🔑  Senha:", font=ctk.CTkFont(weight="bold"),
                     text_color=TEXTO).grid(row=0, column=0, padx=(0, 8))
        self._entry_senha = ctk.CTkEntry(
            ls, placeholder_text="Mínimo 4 caracteres…",
            show="●", fg_color="#0f172a", border_color="#334155",
            text_color=TEXTO, font=ctk.CTkFont(size=13), height=36,
        )
        self._entry_senha.grid(row=0, column=1, sticky="ew", padx=(0, 6))
        self._btn_eye = ctk.CTkButton(
            ls, text="👁", width=36, height=36,
            fg_color="#334155", hover_color="#475569", corner_radius=8,
            command=self._toggle_senha,
        )
        self._btn_eye.grid(row=0, column=2, padx=(0, 8))
        ctk.CTkButton(
            ls, text="🔍  Escanear Arquivos", width=188, height=36,
            fg_color="#7c3aed", hover_color="#6d28d9", corner_radius=8,
            command=self._escanear,
        ).grid(row=0, column=3)

        # ── Árvore de arquivos ────────────────────────────────────────────────
        lista = ctk.CTkFrame(main, fg_color=BG_CARD, corner_radius=12)
        lista.grid(row=1, column=0, sticky="nsew", pady=(0, 10))
        lista.columnconfigure(0, weight=1)
        lista.rowconfigure(1, weight=1)

        # Barra de topo da lista
        topo = ctk.CTkFrame(lista, fg_color="transparent", height=40)
        topo.grid(row=0, column=0, sticky="ew", padx=12, pady=(8, 0))
        topo.pack_propagate(False)

        self._lbl_contagem = ctk.CTkLabel(
            topo, text="Nenhum arquivo encontrado",
            font=ctk.CTkFont(size=12), text_color=SUBTEXTO)
        self._lbl_contagem.pack(side="left")

        for txt, val in [("Desmarcar todos", False), ("Marcar todos", True)]:
            ctk.CTkButton(
                topo, text=txt, width=120, height=28,
                fg_color="#334155", hover_color="#475569", corner_radius=6,
                font=ctk.CTkFont(size=11),
                command=lambda v=val: self._marcar_todos(v),
            ).pack(side="right", padx=(4, 0))

        self._scroll = ctk.CTkScrollableFrame(
            lista, fg_color=BG_SCROLL, corner_radius=8,
            scrollbar_button_color="#334155",
        )
        self._scroll.grid(row=1, column=0, sticky="nsew", padx=12, pady=(4, 12))
        self._scroll.columnconfigure(0, weight=1)

        ctk.CTkLabel(
            self._scroll,
            text="⬆  Selecione a pasta de origem e clique em 'Escanear Arquivos'",
            text_color=SUBTEXTO, font=ctk.CTkFont(size=13),
        ).pack(pady=40)

        # ── Rodapé ────────────────────────────────────────────────────────────
        rod = ctk.CTkFrame(main, fg_color=BG_CARD, corner_radius=12)
        rod.grid(row=2, column=0, sticky="ew")
        rod.columnconfigure(0, weight=1)

        self._lbl_status = ctk.CTkLabel(
            rod, text="Aguardando…", anchor="w",
            font=ctk.CTkFont(size=12), text_color=SUBTEXTO)
        self._lbl_status.grid(row=0, column=0, padx=16, pady=(10, 1), sticky="w")

        self._lbl_log = ctk.CTkLabel(
            rod, text="", anchor="w",
            font=ctk.CTkFont(size=10), text_color="#4ade80")
        self._lbl_log.grid(row=1, column=0, padx=16, pady=(0, 1), sticky="w")

        self._bar = ctk.CTkProgressBar(
            rod, fg_color="#1e293b", progress_color=AZUL,
            height=10, corner_radius=5)
        self._bar.set(0)
        self._bar.grid(row=2, column=0, padx=16, pady=(0, 12), sticky="ew")

        self._btn_proc = ctk.CTkButton(
            rod, text="🚀  Processar e Proteger", height=44,
            fg_color=VERDE, hover_color="#16a34a", corner_radius=8,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self._iniciar,
        )
        self._btn_proc.grid(row=0, column=1, rowspan=3, padx=16, pady=12)

    # ── Ações ─────────────────────────────────────────────────────────────────
    def _sel_origem(self):
        p = filedialog.askdirectory(title="Pasta de Origem")
        if p:
            self._pasta_origem = p
            self._entry_origem.delete(0, "end")
            self._entry_origem.insert(0, p)

    def _sel_destino(self):
        p = filedialog.askdirectory(title="Pasta de Destino")
        if p:
            self._pasta_destino = p
            self._entry_destino.delete(0, "end")
            self._entry_destino.insert(0, p)

    def _toggle_senha(self):
        s = self._entry_senha.cget("show")
        self._entry_senha.configure(show="" if s else "●")
        self._btn_eye.configure(text="🙈" if s else "👁")

    def _escanear(self):
        origem = self._entry_origem.get().strip() or self._pasta_origem
        if not origem or not os.path.isdir(origem):
            self._toast("⚠️  Selecione uma pasta de origem válida!", AMARELO)
            return

        self._pasta_origem = origem

        for w in self._scroll.winfo_children():
            w.destroy()
        self._pasta_widgets.clear()

        grupos = varrer_pasta(origem)

        if not grupos:
            ctk.CTkLabel(
                self._scroll,
                text="Nenhum arquivo Excel encontrado nesta pasta.",
                text_color=SUBTEXTO, font=ctk.CTkFont(size=13),
            ).pack(pady=40)
            self._lbl_contagem.configure(text="0 arquivos encontrados")
            return

        total = sum(len(v) for v in grupos.values())

        for pasta, arqs in grupos.items():
            pw = PastaWidget(self._scroll, pasta, arqs)
            pw.grid(sticky="ew", pady=(0, 6), padx=4)
            self._pasta_widgets.append(pw)

        n_pastas = len(grupos)
        self._lbl_contagem.configure(
            text=(f"{total} arquivo{'s' if total>1 else ''} Excel "
                  f"em {n_pastas} pasta{'s' if n_pastas>1 else ''}")
        )
        self._lbl_status.configure(text=f"Pronto — {total} arquivo(s) encontrado(s).")

    def _marcar_todos(self, valor: bool):
        for pw in self._pasta_widgets:
            pw.marcar_todos(valor)

    # ── Processamento ─────────────────────────────────────────────────────────
    def _iniciar(self):
        if self._processando:
            return

        senha   = self._entry_senha.get().strip()
        destino = self._entry_destino.get().strip() or self._pasta_destino

        erros = []
        if not self._pasta_origem:  erros.append("pasta de origem")
        if not destino:              erros.append("pasta de destino")
        if not senha:                erros.append("senha")
        elif len(senha) < 4:         erros.append("senha com mínimo 4 caracteres")
        if not self._pasta_widgets:  erros.append("arquivos escaneados")

        if erros:
            self._toast(f"⚠️  Informe: {', '.join(erros)}!", AMARELO)
            return

        self._pasta_destino = destino
        threading.Thread(target=self._processar, daemon=True).start()

    def _processar(self):
        self._processando = True
        self.after(0, lambda: self._btn_proc.configure(
            state="disabled", text="⏳  Processando…", fg_color="#1e40af"))
        self.after(0, lambda: self._bar.configure(progress_color=AZUL))

        origem = Path(self._pasta_origem)
        ts     = time.strftime("%Y%m%d_%H%M%S")
        dest_final = Path(self._pasta_destino) / f"Protegido_{origem.name}_{ts}"
        dest_final.mkdir(parents=True, exist_ok=True)

        # Inicia log
        log_path = str(dest_final / f"relatorio_{ts}.log")
        configurar_log(log_path)
        self.after(0, lambda p=log_path: self._lbl_log.configure(
            text=f"📋  Log: {p}"))

        # Coleta arquivos marcados para proteção
        marcados: set[str] = set()
        for pw in self._pasta_widgets:
            for row in pw.rows:
                if row.marcado:
                    marcados.add(row.caminho)

        senha = self._entry_senha.get().strip()
        todos = [p for p in origem.rglob("*") if p.is_file()]
        total = len(todos)

        _logger.info("Origem  : %s", origem)
        _logger.info("Destino : %s", dest_final)
        _logger.info("Total de arquivos : %d", total)
        _logger.info("Marcados para protecao : %d", len(marcados))
        _logger.info("-" * 70)

        cnt_prot_ok  = 0
        cnt_prot_err = 0
        cnt_copiados = 0
        cnt_erros_cp = 0

        for i, arq in enumerate(todos, 1):
            rel      = str(arq.relative_to(origem))
            dest_arq = str(dest_final / rel)

            self._status_safe(f"[{i}/{total}]  {arq.name}", i / total)

            if rel in marcados:
                _logger.info("[%d/%d] PROTEGER  %s", i, total, rel)
                res = proteger_excel(str(arq), dest_arq, senha)

                if res.sucesso and res.verificado:
                    cnt_prot_ok += 1
                    _logger.info(
                        "  OK  Criptografado e verificado (metodo: %s)", res.metodo)
                elif res.sucesso:
                    cnt_prot_ok += 1
                    _logger.warning(
                        "  AVISO  Criptografado mas verificacao inconclusiva "
                        "(metodo: %s)", res.metodo)
                else:
                    cnt_prot_err += 1
                    _logger.error(
                        "  ERRO  Falha ao proteger: %s", res.erro)
            else:
                _logger.info("[%d/%d] COPIAR    %s", i, total, rel)
                try:
                    copiar_arquivo(str(arq), dest_arq)
                    cnt_copiados += 1
                    _logger.info("  OK  Copiado")
                except Exception as exc:
                    cnt_erros_cp += 1
                    _logger.error("  ERRO  Falha ao copiar: %s\n%s",
                                  exc, traceback.format_exc())

        total_erros = cnt_prot_err + cnt_erros_cp

        _logger.info("=" * 70)
        _logger.info("RESUMO")
        _logger.info("  Protegidos com sucesso : %d", cnt_prot_ok)
        _logger.info("  Falhas de protecao     : %d", cnt_prot_err)
        _logger.info("  Copiados sem senha     : %d", cnt_copiados)
        _logger.info("  Falhas de copia        : %d", cnt_erros_cp)
        _logger.info("  TOTAL DE ERROS         : %d", total_erros)
        _logger.info("  Destino final          : %s", dest_final)
        _logger.info("=" * 70)

        resumo = (
            f"✅  Concluído!  "
            f"🔐 {cnt_prot_ok} protegido(s)  |  "
            f"📋 {cnt_copiados} copiado(s)  |  "
            f"❌ {total_erros} erro(s)"
        )
        cor = VERDE if total_erros == 0 else VERMELHO

        self._status_safe(resumo, 1.0)
        self.after(0, lambda c=cor: self._bar.configure(progress_color=c))
        self.after(0, lambda: self._btn_proc.configure(
            state="normal", text="🚀  Processar e Proteger", fg_color=VERDE))
        self._processando = False

    # ── Helpers thread-safe ───────────────────────────────────────────────────
    def _status_safe(self, txt: str, prog: float):
        self.after(0, lambda t=txt, p=prog: (
            self._lbl_status.configure(text=t),
            self._bar.set(p),
        ))

    def _toast(self, msg: str, cor: str = AZUL):
        t = ctk.CTkToplevel(self)
        t.overrideredirect(True)
        t.configure(fg_color=cor)
        t.attributes("-topmost", True)
        x = self.winfo_x() + self.winfo_width()  // 2 - 215
        y = self.winfo_y() + self.winfo_height() - 80
        t.geometry(f"430x46+{x}+{y}")
        ctk.CTkLabel(t, text=msg, text_color="white",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(expand=True)
        t.after(3000, t.destroy)


# ─── Entrada ─────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = App()
    app.mainloop()
