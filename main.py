import os
import sys
import math
import logging
import subprocess
from datetime import datetime
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
from PIL import Image, ImageDraw


# =============================================================================
# LOGGING
# =============================================================================
logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] %(levelname)s: %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler("app.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)


# =============================================================================
# PALETA — macOS Sonoma Light/Dark adaptável
# =============================================================================
COLORS = {
    # Fundos
    "bg_window": ("#F2F2F7", "#1C1C1E"),
    "bg_sidebar": ("#FFFFFF", "#2C2C2E"),
    "bg_header": ("#F2F2F7", "#1C1C1E"),
    "bg_item_hover": ("#E5E5EA", "#3A3A3C"),
    "bg_folder_root": ("#FFFFFF", "#2C2C2E"),
    # Botão de script (pill colorido)
    "pill_bg": ("#E8F0FE", "#1C3358"),
    "pill_bg_hover": ("#C7D9FC", "#2C4A7A"),
    "pill_fg": ("#1A56DB", "#6EA8FE"),
    # Textos
    "text_primary": ("#1C1C1E", "#F2F2F7"),
    "text_secondary": ("#6E6E73", "#8E8E93"),
    "text_folder": ("#3A3A3C", "#EBEBF5"),
    # Separador / borda
    "separator": ("#D1D1D6", "#38383A"),
    # Botão chevron
    "chevron": ("#8E8E93", "#636366"),
    # Acento (azul Apple)
    "accent": "#0A84FF",
    "accent_dark": "#0066CC",
}

FONT_TITLE = ("SF Pro Display", "Helvetica Neue", "Arial")
FONT_BODY = ("SF Pro Text", "Helvetica Neue", "Arial")


def _font(family_list, size, weight="normal"):
    for f in family_list:
        try:
            t = ctk.CTkFont(family=f, size=size, weight=weight)
            return t
        except Exception:
            continue
    return ctk.CTkFont(size=size, weight=weight)


# =============================================================================
# LICENSE
# =============================================================================
class LicenseManager:
    @staticmethod
    def verificar_licenca():
        data_expiracao_str = "01/06/2026"
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
    def _alerta(data):
        root = tk.Tk()
        root.withdraw()
        messagebox.showwarning("Licença Expirada", f"Acesso encerrado em {data}.")
        root.destroy()
        sys.exit(0)


# =============================================================================
# ICON FACTORY — gear sólido via PIL, adaptado a light/dark
# =============================================================================
class IconFactory:
    """
    Gera ícones vetoriais renderizados com PIL.
    Retorna CTkImage (light + dark) pronto para usar em CTkButton/CTkLabel.
    """

    @staticmethod
    def _draw_gear(size: int, fg: str, bg: str) -> Image.Image:
        """
        Desenha um gear sólido num canvas RGBA de `size` × `size` px.
        Usa supersampling 4× para bordas suaves e depois reduz.
        """
        S = size * 4  # supersampling 4×
        cx = cy = S // 2
        r_outer = S * 0.42  # raio externo dos dentes
        r_inner = S * 0.28  # base dos dentes
        r_hub = S * 0.14  # buraco central
        n_teeth = 8
        tooth_w = 0.30  # largura angular do dente (rad fração)

        img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        # ── corpo do gear (polígono com dentes) ──────────────────────────────
        pts = []
        step = math.tau / n_teeth
        for i in range(n_teeth):
            base = step * i
            # flanco esquerdo da base
            a0 = base - step * 0.22
            pts.append((cx + r_inner * math.cos(a0), cy + r_inner * math.sin(a0)))
            # sobe ao dente
            a1 = base - step * tooth_w
            pts.append((cx + r_outer * math.cos(a1), cy + r_outer * math.sin(a1)))
            # topo do dente
            a2 = base + step * tooth_w
            pts.append((cx + r_outer * math.cos(a2), cy + r_outer * math.sin(a2)))
            # desce da base
            a3 = base + step * 0.22
            pts.append((cx + r_inner * math.cos(a3), cy + r_inner * math.sin(a3)))

        draw.polygon(pts, fill=fg)

        # ── buraco central (RGBA punch-through) ─────────────────────────────
        cx_f = cy_f = S / 2
        draw.ellipse(
            (cx_f - r_hub, cy_f - r_hub, cx_f + r_hub, cy_f + r_hub),
            fill=(0, 0, 0, 0),
        )

        # reduz para tamanho final com antialiasing
        return img.resize((size, size), Image.LANCZOS)

    @staticmethod
    def _hex_to_rgba(h: str) -> tuple:
        h = h.lstrip("#")
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        return (r, g, b, 255)

    @classmethod
    def gear(
        cls, size: int = 18, fg_light: str = "#1A56DB", fg_dark: str = "#6EA8FE"
    ) -> ctk.CTkImage:
        """
        Retorna CTkImage com versão light e dark do gear.
        O fundo é sempre transparente — combina com qualquer bg de botão.
        """
        # Converte hex → RGBA para PIL
        light_img = cls._draw_gear(size, cls._hex_to_rgba(fg_light), None)
        dark_img = cls._draw_gear(size, cls._hex_to_rgba(fg_dark), None)
        return ctk.CTkImage(
            light_image=light_img, dark_image=dark_img, size=(size, size)
        )

    # Cache de instância — evita recriar a mesma imagem a cada ScriptItem
    _cache: dict = {}

    @classmethod
    def get(cls, key: str = "gear", **kwargs) -> ctk.CTkImage:
        if key not in cls._cache:
            cls._cache[key] = cls.gear(**kwargs)
        return cls._cache[key]


# =============================================================================
# SCRIPT ITEM — pill colorido, compacto
# =============================================================================
class ScriptItem(ctk.CTkFrame):
    def __init__(self, master, nome, caminho, nivel=0, on_run=None, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.grid_columnconfigure(0, weight=1)

        indent = nivel * 16

        # Container do pill
        row = ctk.CTkFrame(self, fg_color="transparent")
        row.grid(row=0, column=0, sticky="ew", padx=(indent + 8, 8), pady=1)
        row.grid_columnconfigure(0, weight=1)

        label_nome = nome[:-3] if nome.endswith(".py") else nome  # remove .py do label

        # Ícone gear sólido (PIL) — light #1A56DB / dark #6EA8FE
        _icon = IconFactory.get(
            "gear",
            size=15,
            fg_light=COLORS["pill_fg"][0],
            fg_dark=COLORS["pill_fg"][1],
        )

        self.btn = ctk.CTkButton(
            row,
            text=f" {label_nome}",
            image=_icon,
            compound="left",  # ícone à esquerda do texto
            anchor="w",
            height=28,
            corner_radius=7,
            border_width=0,
            fg_color=COLORS["pill_bg"],
            hover_color=COLORS["pill_bg_hover"],
            text_color=COLORS["pill_fg"],
            font=_font(FONT_BODY, 12),
            command=lambda: on_run(caminho, nome) if on_run else None,
        )
        self.btn.grid(row=0, column=0, sticky="ew")


# =============================================================================
# FOLDER NODE — colapsável, estilo Finder sidebar
# =============================================================================
class FolderNode(ctk.CTkFrame):
    def __init__(self, master, nome, nivel=0, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.nivel = nivel
        self.expanded = True
        self.grid_columnconfigure(0, weight=1)

        indent = nivel * 16

        # ── Cabeçalho da pasta ──────────────────────────────────────────────
        self.header_row = ctk.CTkFrame(
            self,
            fg_color=COLORS["bg_folder_root"] if nivel == 0 else "transparent",
            corner_radius=8 if nivel == 0 else 0,
        )
        self.header_row.grid(
            row=0,
            column=0,
            sticky="ew",
            padx=(indent, 8) if nivel > 0 else (8, 8),
            pady=(4, 0) if nivel == 0 else (2, 0),
        )
        self.header_row.grid_columnconfigure(1, weight=1)

        # Chevron
        self.lbl_chevron = ctk.CTkLabel(
            self.header_row,
            text="▾",
            width=18,
            font=_font(FONT_BODY, 11),
            text_color=COLORS["chevron"],
            cursor="hand2",
        )
        self.lbl_chevron.grid(row=0, column=0, padx=(8, 2), pady=4)
        self.lbl_chevron.bind("<Button-1>", lambda e: self.toggle())

        # Ícone + nome
        icon = "📂" if nivel == 0 else "📁"
        self.lbl_nome = ctk.CTkLabel(
            self.header_row,
            text=f"{icon}  {nome}",
            anchor="w",
            font=_font(FONT_BODY, 12, "bold") if nivel == 0 else _font(FONT_BODY, 11),
            text_color=COLORS["text_folder"],
            cursor="hand2",
        )
        self.lbl_nome.grid(row=0, column=1, sticky="ew", padx=(0, 8), pady=4)
        self.lbl_nome.bind("<Button-1>", lambda e: self.toggle())

        # Contador (populado depois)
        self.lbl_count = ctk.CTkLabel(
            self.header_row,
            text="",
            width=28,
            font=_font(FONT_BODY, 10),
            text_color=COLORS["text_secondary"],
        )
        self.lbl_count.grid(row=0, column=2, padx=(0, 8))

        # Separador sutil abaixo do cabeçalho raiz
        if nivel == 0:
            sep = ctk.CTkFrame(self, fg_color=COLORS["separator"], height=1)
            sep.grid(row=1, column=0, sticky="ew", padx=16, pady=(2, 0))

        # ── Container de conteúdo ────────────────────────────────────────────
        self.container = ctk.CTkFrame(self, fg_color="transparent")
        self.container.grid(row=2, column=0, sticky="ew", padx=0, pady=(0, 4))
        self.container.grid_columnconfigure(0, weight=1)

        self._script_count = 0

    # ── Colapsar / Expandir ─────────────────────────────────────────────────
    def toggle(self):
        self.expanded = not self.expanded
        if self.expanded:
            self.container.grid()
            self.lbl_chevron.configure(text="▾")
        else:
            self.container.grid_remove()
            self.lbl_chevron.configure(text="▸")

    def set_count(self, n):
        self._script_count = n
        if n > 0:
            self.lbl_count.configure(text=f"{n}")


# =============================================================================
# APP PRINCIPAL
# =============================================================================
class LauncherApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Prometeus Ecosystem")
        self.geometry("560x760")
        self.minsize(480, 480)
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.diretorio_base = os.path.dirname(os.path.abspath(__file__))
        self.target_folders = ["App.Gemco", "Ar.Excel", "NF_CTE", "Utilitários"]

        self._total_scripts = 0

        self._setup_ui()
        self._build_tree()
        self._update_subtitle()

    # ── UI ──────────────────────────────────────────────────────────────────
    def _setup_ui(self):
        self.configure(fg_color=COLORS["bg_window"])
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # ── Header ──────────────────────────────────────────────────────────
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=24, pady=(28, 0))
        header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            header,
            text="Prometeus Ecosystem",
            font=_font(FONT_TITLE, 22, "bold"),
            text_color=COLORS["text_primary"],
            anchor="w",
        ).grid(row=0, column=0, sticky="w")

        self.lbl_subtitle = ctk.CTkLabel(
            header,
            text="Carregando…",
            font=_font(FONT_BODY, 12),
            text_color=COLORS["text_secondary"],
            anchor="w",
        )
        self.lbl_subtitle.grid(row=1, column=0, sticky="w", pady=(2, 0))

        # Botão collapse all / expand all
        ctrl_row = ctk.CTkFrame(self, fg_color="transparent")
        ctrl_row.grid(row=0, column=0, sticky="ne", padx=24, pady=(32, 0))

        self.btn_collapse_all = ctk.CTkButton(
            ctrl_row,
            text="Recolher tudo",
            width=110,
            height=26,
            corner_radius=6,
            font=_font(FONT_BODY, 11),
            fg_color="transparent",
            border_width=1,
            border_color=COLORS["separator"],
            text_color=COLORS["text_secondary"],
            hover_color=COLORS["bg_item_hover"],
            command=self._collapse_all,
        )
        self.btn_collapse_all.pack(side="left", padx=(0, 6))

        self.btn_expand_all = ctk.CTkButton(
            ctrl_row,
            text="Expandir tudo",
            width=100,
            height=26,
            corner_radius=6,
            font=_font(FONT_BODY, 11),
            fg_color="transparent",
            border_width=1,
            border_color=COLORS["separator"],
            text_color=COLORS["text_secondary"],
            hover_color=COLORS["bg_item_hover"],
            command=self._expand_all,
        )
        self.btn_expand_all.pack(side="left")

        # ── Linha divisória ──────────────────────────────────────────────────
        ctk.CTkFrame(self, fg_color=COLORS["separator"], height=1).grid(
            row=0, column=0, sticky="sew", padx=20, pady=(72, 0)
        )

        # ── Área scrollável da árvore ────────────────────────────────────────
        self.tree_scroll = ctk.CTkScrollableFrame(
            self,
            fg_color=COLORS["bg_window"],
            scrollbar_button_color=COLORS["separator"],
            scrollbar_button_hover_color=COLORS["chevron"],
            corner_radius=0,
        )
        self.tree_scroll.grid(row=1, column=0, sticky="nsew", padx=0, pady=(8, 0))
        self.tree_scroll.grid_columnconfigure(0, weight=1)

        # ── Footer ───────────────────────────────────────────────────────────
        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.grid(row=2, column=0, sticky="ew", padx=24, pady=(6, 14))
        footer.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            footer,
            text="© Roberto Santos",
            font=_font(FONT_BODY, 10),
            text_color=COLORS["text_secondary"],
            anchor="w",
        ).grid(row=0, column=0, sticky="w")

    # ── Construção da árvore ────────────────────────────────────────────────
    def _build_tree(self):
        self._folder_nodes: list[FolderNode] = []
        encontrou = False

        for pasta_raiz in self.target_folders:
            caminho_raiz = os.path.join(self.diretorio_base, pasta_raiz)
            if not os.path.exists(caminho_raiz):
                logging.warning(f"Diretório não encontrado: {pasta_raiz}")
                continue

            encontrou = True
            no_raiz = FolderNode(self.tree_scroll, pasta_raiz, nivel=0)
            no_raiz.grid(
                row=len(self._folder_nodes), column=0, sticky="ew", padx=8, pady=(4, 0)
            )
            self._folder_nodes.append(no_raiz)

            contagem_raiz = 0

            for root, dirs, files in os.walk(caminho_raiz):
                dirs.sort()
                nivel_atual = root.replace(caminho_raiz, "").count(os.sep)

                if root != caminho_raiz:
                    nome_sub = os.path.basename(root)
                    no_sub = FolderNode(
                        no_raiz.container, nome_sub, nivel=nivel_atual + 1
                    )
                    no_sub.pack(fill="x", pady=0)
                    self._folder_nodes.append(no_sub)
                    parent_container = no_sub.container
                else:
                    parent_container = no_raiz.container

                scripts = sorted(
                    [f for f in files if f.endswith(".py") and f != "main.py"]
                )
                for f in scripts:
                    full_path = os.path.join(root, f)
                    item = ScriptItem(
                        parent_container,
                        f,
                        full_path,
                        nivel=nivel_atual + 1,
                        on_run=self.executar_script,
                    )
                    item.pack(fill="x")
                    contagem_raiz += 1
                    self._total_scripts += 1
                    logging.info(f"Script: {f} em {root}")

            no_raiz.set_count(contagem_raiz)

        if not encontrou:
            ctk.CTkLabel(
                self.tree_scroll,
                text="Nenhuma pasta de automação encontrada.",
                font=_font(FONT_BODY, 13),
                text_color=COLORS["text_secondary"],
            ).pack(pady=60)

    def _update_subtitle(self):
        n = self._total_scripts
        s = "script" if n == 1 else "scripts"
        f = len(self.target_folders)
        self.lbl_subtitle.configure(
            text=f"{n} {s} disponíve{'l' if n==1 else 'is'} em {f} módulos"
        )

    # ── Colapsar / Expandir todos ────────────────────────────────────────────
    def _collapse_all(self):
        for node in self._folder_nodes:
            if node.expanded:
                node.toggle()

    def _expand_all(self):
        for node in self._folder_nodes:
            if not node.expanded:
                node.toggle()

    # ── Execução ─────────────────────────────────────────────────────────────
    def executar_script(self, caminho, nome):
        logging.info(f"Executando: {nome} | {caminho}")
        try:
            subprocess.Popen([sys.executable, caminho], cwd=os.path.dirname(caminho))
        except Exception as e:
            logging.error(f"Falha: {e}")
            messagebox.showerror(
                "Erro de Execução", f"Não foi possível abrir:\n{nome}\n\n{e}"
            )


# =============================================================================
# ENTRY POINT
# =============================================================================
if __name__ == "__main__":
    if LicenseManager.verificar_licenca():
        try:
            logging.info("Aplicação iniciada.")
            app = LauncherApp()
            app.mainloop()
        except Exception as e:
            logging.critical(f"Erro fatal: {e}")
