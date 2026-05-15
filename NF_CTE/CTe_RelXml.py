from __future__ import annotations

import logging
import os
import queue
import sys
import threading
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from datetime import datetime
from logging.handlers import RotatingFileHandler
from tkinter import filedialog
from typing import Any

import customtkinter as ctk
import pandas as pd


# 1. Configurações de Diretório e Constantes Globais
def obter_diretorio_base() -> str:
    """
    Retorna o diretório raiz da aplicação.
    Compatível com execução direta (.py) e executável compilado via Nuitka.
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


BASE_DIR: str = obter_diretorio_base()

FUNDO_PRINCIPAL = "#0A0A0A"
SUPERFICIE = "#1C1C1C"
BORDA_FORTE = "#2A2A2A"
BORDA_SUTIL = "#3A3A3A"
TEXTO_SECUNDARIO = "#8C8C8C"
TEXTO_PRIMARIO = "#BEBEBE"
TEXTO_DESTAQUE = "#EDEDED"
OURO_PRINCIPAL = "#D4AF37"
ESMERALDA_PRIMARIA = "#00A36C"
ESMERALDA_SUCESSO = "#00C17C"
ERRO = "#C8102E"
AVISO = "#FFB800"

NAMESPACE_CTE = {"ns": "http://www.portalfiscal.inf.br/cte"}


# 2. Configuração de Logs
def criar_logger(nome_modulo: str, usuario: str = "sistema") -> logging.Logger:
    """Cria logger configurado com formato padrão e rotação de arquivo."""
    formato = f"[%(asctime)s],[{usuario}],[{nome_modulo}] %(levelname)s: %(message)s"
    formatador = logging.Formatter(formato, datefmt="%Y-%m-%d %H:%M:%S")

    caminho_log = os.path.join(BASE_DIR, "logs", "cte_rel_xml.log")
    os.makedirs(os.path.dirname(caminho_log), exist_ok=True)

    handler_arquivo = RotatingFileHandler(
        caminho_log, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8"
    )
    handler_arquivo.setFormatter(formatador)

    handler_console = logging.StreamHandler()
    handler_console.setFormatter(formatador)

    logger = logging.getLogger(nome_modulo)
    logger.setLevel(logging.INFO)
    logger.addHandler(handler_arquivo)
    logger.addHandler(handler_console)
    return logger


logger = criar_logger("cte_rel_xml")


# 3. Exceções Customizadas e Modelos
class ErroProcessamentoCTe(Exception):
    """Levantado quando ocorre um erro na extração de dados do CT-e."""


@dataclass
class CTeDados:
    """Estrutura de dados para armazenar informações extraídas de um CT-e."""
    n_cte: str = ""
    data: str = ""
    cnpj_emit: str = ""
    emitente: str = ""
    cnpj_remetente: str = ""
    nome_remetente: str = ""
    destinatario: str = ""
    cidade: str = ""
    valor: float = 0.0
    peso_bruto: float = 0.0
    tonelada: float = 0.0
    v_tonelada: float = 0.0
    chave_cte: str = ""

    def para_dicionario(self) -> dict[str, Any]:
        """Converte a classe para dicionário para exportação no Pandas."""
        return {
            "N_CTE": self.n_cte,
            "DATA": self.data,
            "CNPJ_EMIT": self.cnpj_emit,
            "EMITENTE": self.emitente,
            "CNPJ_REMETENTE": self.cnpj_remetente,
            "NOME_REMETENTE": self.nome_remetente,
            "DESTINATARIO": self.destinatario,
            "CIDADE": self.cidade,
            "VALOR": self.valor,
            "PESO_BRUTO": self.peso_bruto,
            "TONELADA": self.tonelada,
            "V_TONELADA": self.v_tonelada,
            "CHAVE_CTE": self.chave_cte,
        }


# 4. Serviços (Lógica de Negócio)
class ServicoProcessamentoCTe:
    """Processa arquivos XML de CT-e e exporta para Excel."""

    def _obter_texto_tag(self, raiz: ET.Element, caminho: str, padrao: str = "") -> str:
        """Busca o texto de uma tag XML segura contra ausências."""
        elemento = raiz.find(caminho, NAMESPACE_CTE)
        if elemento is not None and elemento.text is not None:
            return elemento.text
        return padrao

    def _formatar_data(self, data_bruta: str) -> str:
        """Formata a data de emissão para o padrão DD/MM/YYYY."""
        if not data_bruta:
            return ""
        try:
            parte_data = data_bruta[:10]
            return datetime.strptime(parte_data, "%Y-%m-%d").strftime("%d/%m/%Y")
        except ValueError:
            logger.warning("ServicoProcessamentoCTe | falha formatar data | valor=%s", data_bruta)
            return data_bruta

    def processar_arquivo(self, caminho_arquivo: str) -> CTeDados:
        """Extrai dados de um único XML e os retorna num modelo consolidado."""
        try:
            arvore = ET.parse(caminho_arquivo)
            raiz = arvore.getroot()

            valor_txt = self._obter_texto_tag(raiz, ".//ns:vPrest/ns:vTPrest", "0")
            peso_txt = self._obter_texto_tag(raiz, ".//ns:infCarga/ns:infQ/ns:qCarga", "0")

            valor_prest = float(valor_txt.replace(",", "."))
            peso_bruto = float(peso_txt.replace(",", "."))
            tonelada = peso_bruto / 1000
            v_tonelada = valor_prest / tonelada if tonelada > 0 else 0.0

            chave_raw = self._obter_texto_tag(raiz, ".//ns:chCTe")
            if not chave_raw:
                inf_cte = raiz.find(".//ns:infCte", NAMESPACE_CTE)
                chave_raw = inf_cte.attrib.get("Id", "")[3:] if inf_cte is not None else ""

            destinatario = self._obter_texto_tag(raiz, ".//ns:dest/ns:CNPJ") or self._obter_texto_tag(
                raiz, ".//ns:dest/ns:CPF"
            )

            return CTeDados(
                n_cte=self._obter_texto_tag(raiz, ".//ns:nCT"),
                data=self._formatar_data(self._obter_texto_tag(raiz, ".//ns:dhEmi")),
                cnpj_emit=self._obter_texto_tag(raiz, ".//ns:emit/ns:CNPJ"),
                emitente=self._obter_texto_tag(raiz, ".//ns:emit/ns:xNome"),
                cnpj_remetente=self._obter_texto_tag(raiz, ".//ns:rem/ns:CNPJ"),
                nome_remetente=self._obter_texto_tag(raiz, ".//ns:rem/ns:xNome"),
                destinatario=destinatario,
                cidade=self._obter_texto_tag(raiz, ".//ns:dest/ns:enderDest/ns:xMun"),
                valor=valor_prest,
                peso_bruto=peso_bruto,
                tonelada=tonelada,
                v_tonelada=v_tonelada,
                chave_cte=f"'{chave_raw}",
            )
        except Exception as erro:
            logger.exception("ServicoProcessamentoCTe | erro em arquivo | arquivo=%s", caminho_arquivo)
            raise ErroProcessamentoCTe(f"Falha no arquivo {os.path.basename(caminho_arquivo)}") from erro

    def exportar_para_excel(self, dados: list[CTeDados], pasta_destino: str) -> str:
        """Exporta os dados extraídos para um arquivo Excel na pasta destino."""
        if not dados:
            raise ErroProcessamentoCTe("Nenhum dado para exportar.")

        try:
            lista_dicionarios = [item.para_dicionario() for item in dados]
            df = pd.DataFrame(lista_dicionarios)
            nome_arquivo = f"Relatorio_CTEs_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            caminho_final = os.path.join(pasta_destino, nome_arquivo)

            df.to_excel(caminho_final, index=False)
            return caminho_final
        except Exception as erro:
            logger.exception("ServicoProcessamentoCTe | erro excel | destino=%s", pasta_destino)
            raise ErroProcessamentoCTe("Falha ao salvar Excel. Verifique se o arquivo está aberto.") from erro


# 5. Interface Gráfica (UI)
ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")


class AplicacaoCTe(ctk.CTk):
    """Interface principal da aplicação de extração de CT-e."""

    def __init__(self) -> None:
        super().__init__()

        self.title("Extrator de XML CT-e - v1.4")
        self.geometry("450x600")
        self.resizable(False, False)
        self.configure(fg_color=FUNDO_PRINCIPAL)

        self._servico = ServicoProcessamentoCTe()
        self._fila_ui: queue.Queue[dict[str, Any]] = queue.Queue()

        self._pasta_origem: ctk.StringVar = ctk.StringVar()
        self._pasta_destino: ctk.StringVar = ctk.StringVar()

        self._configurar_grid()
        self._construir_interface()
        self._iniciar_loop_fila()

        logger.info("AplicacaoCTe | inicializada com sucesso")

    def _configurar_grid(self) -> None:
        """Configura a malha principal da janela."""
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

    def _construir_interface(self) -> None:
        """Orquestra a criação dos painéis da interface."""
        self._construir_cabecalho()
        self._construir_conteudo()
        self._construir_rodape()

    def _construir_cabecalho(self) -> None:
        """Cria o topo da aplicação com o título."""
        frame_cabecalho = ctk.CTkFrame(self, fg_color="transparent")
        frame_cabecalho.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 10))
        titulo = ctk.CTkLabel(
            frame_cabecalho,
            text="Processador de CT-e",
            font=ctk.CTkFont(family="Inter", size=24, weight="bold"),
            text_color=OURO_PRINCIPAL,
        )
        titulo.pack()

    def _construir_rodape(self) -> None:
        """Cria o rodapé de assinatura fixado na parte inferior."""
        rodape = ctk.CTkLabel(
            self,
            text="Roberto Santos [LABS]©",
            font=ctk.CTkFont(size=10),
            text_color=TEXTO_SECUNDARIO,
        )
        rodape.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 8))

    def _construir_conteudo(self) -> None:
        """Cria o card central que hospeda os inputs e controles."""
        frame_conteudo = ctk.CTkFrame(self, fg_color=SUPERFICIE, corner_radius=15)
        frame_conteudo.grid(row=1, column=0, sticky="nsew", padx=20, pady=10)

        self._construir_campos_origem(frame_conteudo)
        self._construir_campos_destino(frame_conteudo)
        self._construir_area_progresso(frame_conteudo)

    def _construir_campos_origem(self, parent: ctk.CTkFrame) -> None:
        """Controles para selecionar a pasta contendo os XMLs."""
        lbl = ctk.CTkLabel(parent, text="Pasta dos XMLs Origem", text_color=TEXTO_SECUNDARIO)
        lbl.pack(anchor="w", padx=20, pady=(20, 5))
        
        frame_input = ctk.CTkFrame(parent, fg_color="transparent")
        frame_input.pack(fill="x", padx=20)
        
        entrada = ctk.CTkEntry(
            frame_input,
            textvariable=self._pasta_origem,
            fg_color=FUNDO_PRINCIPAL,
            border_color=BORDA_FORTE,
            state="readonly",
        )
        entrada.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        btn = ctk.CTkButton(
            frame_input,
            text="Buscar",
            width=80,
            fg_color=BORDA_FORTE,
            hover_color=BORDA_SUTIL,
            command=self._selecionar_origem,
        )
        btn.pack(side="right")

    def _construir_campos_destino(self, parent: ctk.CTkFrame) -> None:
        """Controles para selecionar onde o arquivo Excel será salvo."""
        lbl = ctk.CTkLabel(parent, text="Pasta Destino (Excel)", text_color=TEXTO_SECUNDARIO)
        lbl.pack(anchor="w", padx=20, pady=(20, 5))
        
        frame_input = ctk.CTkFrame(parent, fg_color="transparent")
        frame_input.pack(fill="x", padx=20)
        
        entrada = ctk.CTkEntry(
            frame_input,
            textvariable=self._pasta_destino,
            fg_color=FUNDO_PRINCIPAL,
            border_color=BORDA_FORTE,
            state="readonly",
        )
        entrada.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        btn = ctk.CTkButton(
            frame_input,
            text="Buscar",
            width=80,
            fg_color=BORDA_FORTE,
            hover_color=BORDA_SUTIL,
            command=self._selecionar_destino,
        )
        btn.pack(side="right")

    def _construir_area_progresso(self, parent: ctk.CTkFrame) -> None:
        """Cria barra de progresso, status e botão de processamento."""
        self._barra_progresso = ctk.CTkProgressBar(parent, progress_color=ESMERALDA_PRIMARIA)
        self._barra_progresso.pack(fill="x", padx=20, pady=(40, 10))
        self._barra_progresso.set(0)

        self._label_status = ctk.CTkLabel(
            parent, text="Aguardando início...", text_color=TEXTO_SECUNDARIO, font=ctk.CTkFont(size=12)
        )
        self._label_status.pack(pady=(0, 20))

        self._btn_iniciar = ctk.CTkButton(
            parent,
            text="INICIAR PROCESSAMENTO",
            font=ctk.CTkFont(weight="bold", size=14),
            height=45,
            fg_color=ESMERALDA_PRIMARIA,
            hover_color=ESMERALDA_SUCESSO,
            command=self._iniciar_processamento,
        )
        self._btn_iniciar.pack(fill="x", padx=20, pady=20)

    def _iniciar_loop_fila(self) -> None:
        """Inicia o processamento contínuo de mensagens da UI na thread principal."""
        self._processar_fila()

    def _processar_fila(self) -> None:
        """Consome mensagens da fila thread-safe e atualiza a interface."""
        try:
            while True:
                mensagem = self._fila_ui.get_nowait()
                self._tratar_mensagem_ui(mensagem)
        except queue.Empty:
            pass
        self.after(50, self._processar_fila)

    def _tratar_mensagem_ui(self, mensagem: dict[str, Any]) -> None:
        """Trata cada tipo de mensagem oriunda da thread de execução."""
        tipo = mensagem.get("tipo")

        if tipo == "progresso":
            self._barra_progresso.set(mensagem.get("valor", 0.0))
            self._label_status.configure(text=mensagem.get("texto", ""), text_color=TEXTO_SECUNDARIO)

        elif tipo == "sucesso":
            self._barra_progresso.set(1.0)
            self._label_status.configure(text=mensagem.get("texto", ""), text_color=ESMERALDA_SUCESSO)
            self._btn_iniciar.configure(state="normal")

        elif tipo == "erro":
            self._barra_progresso.set(0)
            self._label_status.configure(text=mensagem.get("texto", ""), text_color=ERRO)
            self._btn_iniciar.configure(state="normal")

        elif tipo == "aviso":
            self._label_status.configure(text=mensagem.get("texto", ""), text_color=AVISO)
            self._btn_iniciar.configure(state="normal")

    def _selecionar_origem(self) -> None:
        """Abre o diálogo para seleção do diretório com XMLs."""
        caminho = filedialog.askdirectory(title="Selecione a pasta dos XMLs")
        if caminho:
            self._pasta_origem.set(caminho)

    def _selecionar_destino(self) -> None:
        """Abre o diálogo para seleção do diretório onde o Excel será salvo."""
        caminho = filedialog.askdirectory(title="Selecione a pasta para o Excel")
        if caminho:
            self._pasta_destino.set(caminho)

    def _iniciar_processamento(self) -> None:
        """Valida os inputs e delega a operação pesada para background."""
        origem = self._pasta_origem.get()
        destino = self._pasta_destino.get()

        if not origem or not destino:
            self._fila_ui.put({"tipo": "aviso", "texto": "Selecione as pastas primeiro!"})
            return

        self._btn_iniciar.configure(state="disabled")
        self._barra_progresso.set(0)
        self._label_status.configure(text="Iniciando...", text_color=TEXTO_SECUNDARIO)

        self._executar_em_thread(self._executar_processamento_background, origem, destino)

    def _executar_em_thread(self, funcao: Any, *args: Any) -> None:
        """Executa uma função em uma thread isolada para não travar a UI."""
        threading.Thread(target=funcao, args=args, daemon=True).start()

    def _executar_processamento_background(self, origem: str, destino: str) -> None:
        """Lógica executada em background para processar os XMLs e gerar Excel."""
        logger.info("AplicacaoCTe | iniciando lote | origem=%s", origem)
        try:
            arquivos = [arq for arq in os.listdir(origem) if arq.lower().endswith(".xml")]
            total = len(arquivos)

            if total == 0:
                self._fila_ui.put({"tipo": "aviso", "texto": "Nenhum arquivo XML encontrado."})
                return

            dados_extraidos: list[CTeDados] = []

            for idx, arquivo in enumerate(arquivos):
                caminho_arquivo = os.path.join(origem, arquivo)
                try:
                    dados = self._servico.processar_arquivo(caminho_arquivo)
                    dados_extraidos.append(dados)
                except ErroProcessamentoCTe as erro_proc:
                    logger.warning("AplicacaoCTe | xml ignorado | arquivo=%s", arquivo)

                self._fila_ui.put({
                    "tipo": "progresso",
                    "valor": (idx + 1) / total,
                    "texto": f"Processando {idx + 1} de {total}..."
                })

            if dados_extraidos:
                self._fila_ui.put({"tipo": "progresso", "valor": 1.0, "texto": "Gerando Excel..."})
                self._servico.exportar_para_excel(dados_extraidos, destino)
                msg_sucesso = f"Sucesso! {len(dados_extraidos)} registros exportados."
                self._fila_ui.put({"tipo": "sucesso", "texto": msg_sucesso})
            else:
                self._fila_ui.put({"tipo": "aviso", "texto": "Nenhum dado extraído."})

        except Exception as erro:
            logger.exception("AplicacaoCTe | erro fatal na thread | erro=%s", erro)
            self._fila_ui.put({"tipo": "erro", "texto": "Falha na exportação. Verifique os logs."})


if __name__ == "__main__":
    app = AplicacaoCTe()
    app.mainloop()