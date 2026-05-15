from __future__ import annotations

import logging
import os
import subprocess
import sys
import threading
import queue
from datetime import datetime
from logging.handlers import RotatingFileHandler
from typing import Any, Callable, NamedTuple

import customtkinter as ctk
from PIL import Image

# =============================================================================
# CONFIGURAÇÃO DE AMBIENTE E CAMINHOS
# =============================================================================

def obter_diretorio_base() -> str:
    """
    Retorna o diretório raiz da aplicação.

    Compatível com execução direta (.py) e executável compilado via Nuitka.
    Nunca use __file__ diretamente fora desta função.
    """
    if getattr(sys, "frozen", False):
        # Contexto: executável gerado pelo Nuitka
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


BASE_DIR: str = obter_diretorio_base()


def criar_logger(nome_modulo: str, usuario: str = "sistema", nome_log: str | None = None) -> logging.Logger:
    """
    Cria logger configurado com formato padrão e rotação de arquivo.
    Salva logs em PROMETEUS_ROOT_DIR/logs/{nome_log}.log.
    """
    usuario_real = os.environ.get("PROMETEUS_USER", usuario)
    dir_base = os.environ.get("PROMETEUS_ROOT_DIR", BASE_DIR)
    
    # Se nome_log não for fornecido, tenta pegar da env var ou usa 'aplicacao'
    if not nome_log:
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


LOGGER = criar_logger("principal")

# =============================================================================
# CONSTANTES DE DESIGN (PALETA PREMIUM)
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

# =============================================================================
# COMPONENTES DE INTERFACE
# =============================================================================

class ItemDeScript(ctk.CTkFrame):
    """
    Representa um script individual na lista, com botão de execução.
    """

    def __init__(
        self,
        master: Any,
        nome_arquivo: str,
        caminho_completo: str,
        ao_executar: Callable[[str, str], None],
        **kwargs: Any,
    ) -> None:
        super().__init__(master, fg_color="transparent", **kwargs)
        self.caminho_completo = caminho_completo
        self.nome_arquivo = nome_arquivo
        self.ao_executar = ao_executar

        self._configurar_layout()

    def _configurar_layout(self) -> None:
        """Configura a estrutura visual do item."""
        self.grid_columnconfigure(0, weight=1)

        self.container = ctk.CTkFrame(
            self, fg_color=SUPERFICIE, corner_radius=14, border_width=1, border_color=BORDA_FORTE
        )
        self.container.grid(row=0, column=0, sticky="ew", padx=12, pady=6)
        self.container.grid_columnconfigure(0, weight=1)

        self.lbl_nome = ctk.CTkLabel(
            self.container,
            text=f"  📄  {self.nome_arquivo}",
            text_color=TEXTO_PRIMARIO,
            font=ctk.CTkFont(size=13),
            anchor="w",
        )
        self.lbl_nome.grid(row=0, column=0, sticky="w", padx=15, pady=12)

        self.btn_rodar = ctk.CTkButton(
            self.container,
            text="Iniciar",
            width=80,
            height=30,
            fg_color=ESMERALDA_PRIMARIA,
            hover_color=ESMERALDA_SUCESSO,
            text_color=FUNDO_PRINCIPAL,
            corner_radius=10,
            font=ctk.CTkFont(size=12, weight="bold"),
            command=lambda: self.ao_executar(self.caminho_completo, self.nome_arquivo),
        )
        self.btn_rodar.grid(row=0, column=1, padx=15, pady=12)


class NoDePasta(ctk.CTkFrame):
    """
    Representa uma pasta colapsável na árvore de arquivos.
    """

    def __init__(self, master: Any, nome: str, nivel: int = 0, **kwargs: Any) -> None:
        super().__init__(master, fg_color="transparent", **kwargs)
        self.nivel = nivel
        self.expandido = False  # Iniciar recolhido
        self.grid_columnconfigure(0, weight=1)

        self._configurar_cabecalho(nome)
        self._configurar_container_conteudo()

        # Iniciar no estado recolhido
        self.container_filhos.grid_remove()
        self.btn_toggle.configure(text="▸")

    def _configurar_cabecalho(self, nome: str) -> None:
        """Configura a linha de cabeçalho da pasta."""
        indentacao = self.nivel * 18
        bg_header = SUPERFICIE if self.nivel == 0 else "transparent"
        
        self.header = ctk.CTkFrame(
            self, fg_color=bg_header, corner_radius=12 if self.nivel == 0 else 0
        )
        self.header.grid(
            row=0, column=0, sticky="ew", padx=(indentacao + 10, 10), pady=(8 if self.nivel == 0 else 2, 0)
        )
        self.header.grid_columnconfigure(1, weight=1)

        self.btn_toggle = ctk.CTkButton(
            self.header,
            text="▾",
            width=28,
            height=28,
            fg_color="transparent",
            text_color=TEXTO_SECUNDARIO,
            hover_color=BORDA_SUTIL,
            command=self.alternar,
        )
        self.btn_toggle.grid(row=0, column=0, padx=5, pady=5)

        icone = "📂" if self.nivel == 0 else "📁"
        cor_texto = OURO_PRINCIPAL if self.nivel == 0 else TEXTO_PRIMARIO
        peso_fonte = "bold" if self.nivel == 0 else "normal"

        self.lbl_nome = ctk.CTkLabel(
            self.header,
            text=f"{icone}  {nome}",
            text_color=cor_texto,
            font=ctk.CTkFont(size=14, weight=peso_fonte),
        )
        self.lbl_nome.grid(row=0, column=1, sticky="w", padx=5, pady=5)

        self.lbl_qtd = ctk.CTkLabel(self.header, text="", text_color=TEXTO_SECUNDARIO, font=ctk.CTkFont(size=11))
        self.lbl_qtd.grid(row=0, column=2, padx=15)

    def _configurar_container_conteudo(self) -> None:
        """Cria o container para os itens filhos."""
        self.container_filhos = ctk.CTkFrame(self, fg_color="transparent")
        self.container_filhos.grid(row=1, column=0, sticky="ew")
        self.container_filhos.grid_columnconfigure(0, weight=1)

    def alternar(self) -> None:
        """Alterna entre expandido e colapsado."""
        self.expandido = not self.expandido
        if self.expandido:
            self.container_filhos.grid()
            self.btn_toggle.configure(text="▾")
        else:
            self.container_filhos.grid_remove()
            self.btn_toggle.configure(text="▸")

    def atualizar_contagem(self, total: int) -> None:
        """Define o texto de contagem de itens."""
        if total > 0:
            self.lbl_qtd.configure(text=f"{total} itens")

    def definir_estado(self, expandir: bool) -> None:
        """Força um estado específico de expansão."""
        if self.expandido != expandir:
            self.alternar()


# =============================================================================
# JANELA PRINCIPAL
# =============================================================================

class GerenciadorDeAplicacoes(ctk.CTk):
    """
    Aplicação principal para gestão e execução de automações ERP.
    """

    def __init__(self) -> None:
        super().__init__()

        self.title("Prometeus System - ERP Automation")
        self.geometry("450x600")
        self.resizable(False, False)
        
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")
        self.configure(fg_color=FUNDO_PRINCIPAL)

        self.usuario_atual = ""
        self.banco_senhas = {
            "ROBERTO": "rdsf",
            "JOEDSON": "aragao",
            "TAISSA": "fragas",
            "IGOR": "suri",
            "JADSON": "j123",
        }

        self.pastas_alvo = ["App.Gemco", "Ar.Excel", "NF_CTE"]
        self.nos_arvore: list[NoDePasta] = []
        
        # Gerenciamento de Fila de Execução
        self._fila_execucao: queue.Queue[dict[str, str]] = queue.Queue()
        self._trabalhador_ativo: bool = False
        
        self._configurar_grid_base()
        self._exibir_tela_login()

    def _configurar_grid_base(self) -> None:
        """Configura a estrutura de grid da janela principal."""
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.lbl_rodape = ctk.CTkLabel(
            self,
            text="Roberto Santos [LABS]©",
            font=ctk.CTkFont(size=10, weight="bold"),
            text_color=OURO_PRINCIPAL,
            fg_color=SUPERFICIE,
            height=25
        )
        self.lbl_rodape.grid(row=1, column=0, sticky="ew", padx=0, pady=0)

    def _exibir_tela_login(self) -> None:
        """Renderiza a interface de autenticação."""
        self.frame_conteudo = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_conteudo.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
        self.frame_conteudo.grid_rowconfigure((0, 6), weight=1)
        self.frame_conteudo.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            self.frame_conteudo,
            text="PROMETEUS",
            font=ctk.CTkFont(size=36, weight="bold"),
            text_color=OURO_PRINCIPAL
        ).grid(row=1, column=0, pady=(0, 30))

        self.ent_user = self._criar_input("Usuário")
        self.ent_user.grid(row=2, column=0, pady=12)

        self.ent_pass = self._criar_input("Senha", oculto=True)
        self.ent_pass.grid(row=3, column=0, pady=12)

        self.lbl_feedback = ctk.CTkLabel(self.frame_conteudo, text="", text_color=ERRO, font=ctk.CTkFont(size=12))
        self.lbl_feedback.grid(row=4, column=0, pady=5)

        self.btn_login = ctk.CTkButton(
            self.frame_conteudo,
            text="Acessar Sistema",
            command=self._processar_login,
            fg_color=ESMERALDA_PRIMARIA,
            hover_color=ESMERALDA_SUCESSO,
            text_color=FUNDO_PRINCIPAL,
            corner_radius=12,
            height=40,
            width=240,
            font=ctk.CTkFont(weight="bold")
        )
        self.btn_login.grid(row=5, column=0, pady=20)

    def _criar_input(self, placeholder: str, oculto: bool = False) -> ctk.CTkEntry:
        """Helper para criar campos de entrada padronizados."""
        return ctk.CTkEntry(
            self.frame_conteudo,
            placeholder_text=placeholder,
            show="*" if oculto else "",
            width=240,
            height=40,
            corner_radius=12,
            fg_color=SUPERFICIE,
            border_color=BORDA_FORTE,
            text_color=TEXTO_PRIMARIO,
            placeholder_text_color=TEXTO_SECUNDARIO
        )

    def _processar_login(self) -> None:
        """
        Valida as credenciais informadas.
        
        O nome de usuário é tratado como insensível a maiúsculas/minúsculas,
        enquanto a senha mantém a sensibilidade (case-sensitive).
        """
        user = self.ent_user.get().strip().upper()  # Normaliza para comparação
        pwd = self.ent_pass.get().strip()           # Mantém original para senha

        if self.banco_senhas.get(user) == pwd:
            self.usuario_atual = user
            LOGGER.info(f"Login bem-sucedido: {user}")
            self.frame_conteudo.destroy()
            self._exibir_painel_controle()
        else:
            self.lbl_feedback.configure(text="Usuário ou senha inválidos")
            LOGGER.warning(f"Falha de login para o usuário: {user}")

    def _exibir_painel_controle(self) -> None:
        """Renderiza a área principal de scripts."""
        self.frame_conteudo = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_conteudo.grid(row=0, column=0, sticky="nsew", padx=15, pady=15)
        self.frame_conteudo.grid_rowconfigure(1, weight=1)
        self.frame_conteudo.grid_columnconfigure(0, weight=1)

        # Cabeçalho do Painel
        header = ctk.CTkFrame(self.frame_conteudo, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 15))
        header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            header,
            text="Ecosistema de Automação",
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color=TEXTO_DESTAQUE,
            anchor="w"
        ).grid(row=0, column=0, sticky="w")

        ctk.CTkLabel(
            header,
            text=f"Operador: {self.usuario_atual}",
            text_color=TEXTO_SECUNDARIO,
            font=ctk.CTkFont(size=12),
            anchor="w"
        ).grid(row=1, column=0, sticky="w")

        # Botões de Controle Global - Movidos para uma nova linha para evitar sobreposição
        frame_controles = ctk.CTkFrame(header, fg_color="transparent")
        frame_controles.grid(row=2, column=0, sticky="w", pady=(12, 0))

        self.btn_recolher = ctk.CTkButton(
            frame_controles,
            text="Recolher Tudo",
            width=100,
            height=28,
            corner_radius=10,
            fg_color=SUPERFICIE,
            border_width=1,
            border_color=BORDA_FORTE,
            text_color=TEXTO_SECUNDARIO,
            hover_color=BORDA_SUTIL,
            font=ctk.CTkFont(size=11, weight="bold"),
            command=self._recolher_tudo
        )
        self.btn_recolher.grid(row=0, column=0, padx=(0, 8))

        self.btn_expandir = ctk.CTkButton(
            frame_controles,
            text="Mostrar Tudo",
            width=100,
            height=28,
            corner_radius=10,
            fg_color=SUPERFICIE,
            border_width=1,
            border_color=BORDA_FORTE,
            text_color=TEXTO_SECUNDARIO,
            hover_color=BORDA_SUTIL,
            font=ctk.CTkFont(size=11, weight="bold"),
            command=self._expandir_tudo
        )
        self.btn_expandir.grid(row=0, column=1, padx=0)

        # Lista de Scripts (Scrollable)
        self.area_scroll = ctk.CTkScrollableFrame(
            self.frame_conteudo,
            fg_color="transparent",
            scrollbar_button_color=BORDA_FORTE,
            scrollbar_button_hover_color=BORDA_SUTIL,
            corner_radius=0
        )
        self.area_scroll.grid(row=1, column=0, sticky="nsew")
        self.area_scroll.grid_columnconfigure(0, weight=1)

        self._popular_arvore_scripts()

    def _popular_arvore_scripts(self) -> None:
        """Varre os diretórios e reconstrói a árvore de scripts."""
        for idx, nome_pasta in enumerate(self.pastas_alvo):
            caminho = os.path.join(BASE_DIR, nome_pasta)
            
            # Garante existência da pasta
            if not os.path.exists(caminho):
                try:
                    os.makedirs(caminho, exist_ok=True)
                    LOGGER.info(f"Pasta criada: {nome_pasta}")
                except OSError:
                    LOGGER.exception(f"Erro ao criar diretório: {nome_pasta}")
                    continue

            no = NoDePasta(self.area_scroll, nome_pasta, nivel=0)
            no.grid(row=idx, column=0, sticky="ew", pady=4)
            self.nos_arvore.append(no)
            
            self._mapear_subdiretorios(caminho, no)

    def _mapear_subdiretorios(self, raiz: str, no_pai: NoDePasta) -> None:
        """Mapeia recursivamente subpastas e scripts."""
        contador_total = 0
        try:
            for root, dirs, files in os.walk(raiz):
                dirs.sort()
                sub_nivel = root.replace(raiz, "").count(os.sep)
                
                container = no_pai.container_filhos
                if root != raiz:
                    no_sub = NoDePasta(container, os.path.basename(root), nivel=sub_nivel + 1)
                    no_sub.grid(row=contador_total, column=0, sticky="ew")
                    self.nos_arvore.append(no_sub)
                    container = no_sub.container_filhos
                
                is_frozen = getattr(sys, "frozen", False)
                extensoes_alvo = (".exe",) if is_frozen else (".py",)
                scripts = sorted([f for f in files if f.endswith(extensoes_alvo) and not f.startswith("main")])
                for script in scripts:
                    item = ItemDeScript(
                        container, script, os.path.join(root, script), self._executar_async
                    )
                    item.grid(row=contador_total + 500, column=0, sticky="ew")
                    contador_total += 1
                    
            no_pai.atualizar_contagem(contador_total)
        except Exception:
            LOGGER.exception("Falha ao mapear árvore de diretórios")

    def _expandir_tudo(self) -> None:
        """Expande todos os nós da árvore."""
        for no in self.nos_arvore:
            no.definir_estado(True)

    def _recolher_tudo(self) -> None:
        """Recolhe todos os nós da árvore."""
        for no in self.nos_arvore:
            no.definir_estado(False)

    def _executar_async(self, caminho: str, nome: str) -> None:
        """Adiciona o script à fila de execução."""
        LOGGER.info(f"Agendando execução: {nome}")
        self._fila_execucao.put({"caminho": caminho, "nome": nome})
        
        if not self._trabalhador_ativo:
            self._trabalhador_ativo = True
            threading.Thread(target=self._processar_fila_execucao, daemon=True).start()

    def _processar_fila_execucao(self) -> None:
        """Worker thread que processa scripts um a um."""
        while not self._fila_execucao.empty():
            tarefa = self._fila_execucao.get()
            caminho = tarefa["caminho"]
            nome = tarefa["nome"]
            
            self._rotina_execucao_sync(caminho, nome)
            self._fila_execucao.task_done()
            
        self._trabalhador_ativo = False

    def _rotina_execucao_sync(self, caminho: str, nome: str) -> None:
        """Lógica de execução síncrona dentro da thread do worker."""
        # Limpa o nome para o log (ex: AltCust.py -> altcust)
        nome_log = os.path.splitext(nome)[0].lower().replace(" ", "_")
        
        LOGGER.info(f"Iniciando execução sequencial: {nome} | Log: {nome_log}.log")
        try:
            env = os.environ.copy()
            env["PROMETEUS_USER"] = self.usuario_atual or "sistema"
            env["PROMETEUS_ROOT_DIR"] = BASE_DIR
            env["PROMETEUS_APP_NAME"] = nome_log
            env["PROMETEUS_AUTH_TOKEN"] = "PR0M3T3U5_L0CK_2026"

            if getattr(sys, "frozen", False) and caminho.endswith(".exe"):
                cmd = [caminho]
            else:
                cmd = [sys.executable, caminho]
                
            # Usa subprocess.run para esperar a conclusão antes de seguir para o próximo
            resultado = subprocess.run(
                cmd, 
                cwd=os.path.dirname(caminho), 
                env=env,
                capture_output=False  # Saída vai para os logs do próprio sub-app
            )
            
            LOGGER.info(f"Finalizado: {nome} | Status: {resultado.returncode}")
        except Exception:
            LOGGER.exception(f"Falha na execução de {nome}")
            self.after(0, lambda: self._notificar_erro(f"Erro ao executar sub-aplicativo: {nome}"))

    def _notificar_erro(self, msg: str) -> None:
        """Exibe popup de erro na thread principal."""
        from tkinter import messagebox
        messagebox.showerror("Erro de Automação", msg)


def validar_licenca() -> bool:
    """Verifica validade temporal do software."""
    data_limite_str = "15/01/2027"
    try:
        data_limite = datetime.strptime(data_limite_str, "%d/%m/%Y").date()
        if datetime.now().date() <= data_limite:
            return True
        LOGGER.error("Software expirado")
        return False
    except Exception:
        LOGGER.exception("Erro na validação de licença")
        return False


if __name__ == "__main__":
    if validar_licenca():
        try:
            LOGGER.info("Aplicação iniciada com sucesso")
            app = GerenciadorDeAplicacoes()
            app.mainloop()
        except Exception:
            LOGGER.exception("Falha fatal na inicialização da aplicação")
