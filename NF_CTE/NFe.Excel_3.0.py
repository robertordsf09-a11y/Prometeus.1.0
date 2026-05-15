"""
Módulo para conversão de NFe XML (Modelo 55) para Excel.
Refatorado para o padrão premium, com injeção de dependências, logging e thread safety.
"""
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
from typing import Any, Callable

import customtkinter as ctk
import pandas as pd
from tkinter import filedialog, messagebox

# Opcionais se estiverem disponíveis
try:
    import openpyxl
    from openpyxl.styles import numbers
    from openpyxl.utils import get_column_letter
except ImportError:
    pass

# =============================================================================
# CONFIGURAÇÕES E CONSTANTES
# =============================================================================

def obter_diretorio_base() -> str:
    """
    Retorna o diretório raiz da aplicação.
    Compatível com execução direta (.py) e executável compilado via Nuitka.
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE_DIR: str = obter_diretorio_base()

# Paleta de Cores Premium Minimalista
FUNDO_PRINCIPAL: str = "#0A0A0A"
SUPERFICIE: str = "#1C1C1C"
BORDA_FORTE: str = "#2A2A2A"
BORDA_SUTIL: str = "#3A3A3A"
TEXTO_SECUNDARIO: str = "#8C8C8C"
TEXTO_PRIMARIO: str = "#BEBEBE"
TEXTO_DESTAQUE: str = "#EDEDED"
OURO_PRINCIPAL: str = "#D4AF37"
OURO_ESCURO: str = "#B8972E"
ESMERALDA_DEEP: str = "#006D4E"
ESMERALDA_PRIMARIA: str = "#00A36C"
ESMERALDA_SUCESSO: str = "#00C17C"
ERRO: str = "#C8102E"
PERIGO: str = "#8B0000"
AVISO: str = "#FFB800"

PADX_PADRAO: int = 12
PADY_PADRAO: int = 10
RAIO_BORDA: int = 12

# =============================================================================
# LOGGING
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

logger = criar_logger("nfe_excel")

# =============================================================================
# EXCEÇÕES CUSTOMIZADAS
# =============================================================================

class ErroConversao(Exception):
    """Levantado quando falha a conversão de um XML específico."""

# =============================================================================
# SERVIÇOS (LÓGICA DE NEGÓCIO)
# =============================================================================

class ServicoExtracaoNFe:
    """Isola a lógica de extração de dados do XML da NFe."""
    
    def __init__(self, logger: logging.Logger) -> None:
        self._logger = logger
        self._ns = {'ns': 'http://www.portalfiscal.inf.br/nfe'}

    def formatar_valor(self, valor: Any) -> float:
        """Converte strings para float e garante o formato numérico."""
        if valor is None or valor == "":
            return 0.0
        try:
            if isinstance(valor, str):
                valor = valor.replace(',', '.').strip()
            return float(valor)
        except ValueError:
            return 0.0

    def obter_valor_tag(self, elemento: ET.Element | None, nome_tag: str, padrao: str = "") -> str:
        """Busca uma tag recursivamente no elemento."""
        if elemento is None:
            return padrao
        encontrado = elemento.find(f".//ns:{nome_tag}", self._ns)
        if encontrado is not None and encontrado.text:
            return encontrado.text.strip()
        return padrao
    
    def obter_valor_tag_pai(self, elemento_pai: ET.Element | None, nome_tag: str, padrao: str = "") -> str:
        """Busca uma tag especificamente dentro de um elemento pai."""
        if elemento_pai is not None:
            encontrado = elemento_pai.find(f".//ns:{nome_tag}", self._ns)
            if encontrado is not None and encontrado.text:
                return encontrado.text.strip()
        return padrao

    def formatar_data_br(self, data_str: str) -> str:
        """Converte data do formato ISO para DD/MM/YYYY."""
        if not data_str:
            return ""
        try:
            if 'T' in data_str:
                data_str = data_str.split('T')[0]
            if '-' in data_str:
                data_obj = datetime.strptime(data_str, "%Y-%m-%d")
                return data_obj.strftime("%d/%m/%Y")
            return data_str
        except Exception:
            self._logger.warning("Falha ao formatar data: %s", data_str)
            return data_str

    def extrair_dados_gerais(self, root: ET.Element) -> dict[str, Any]:
        """Extrai os dados gerais da Nota Fiscal."""
        n_nf = self.obter_valor_tag(root, "nNF")
        serie = self.obter_valor_tag(root, "serie")
        dh_emi = self.formatar_data_br(self.obter_valor_tag(root, "dhEmi"))
        ch_nfe = self.obter_valor_tag(root, "chNFe")

        emit = root.find(".//ns:emit", self._ns)
        cnpj_emit = self.obter_valor_tag(emit, "CNPJ")
        x_nome_emit = self.obter_valor_tag(emit, "xNome")
        ie_emit = self.obter_valor_tag(emit, "IE")

        dest = root.find(".//ns:dest", self._ns)
        cnpj_dest = self.obter_valor_tag(dest, "CNPJ")
        x_nome_dest = self.obter_valor_tag(dest, "xNome")
        x_mun_dest = self.obter_valor_tag(dest, "xMun")

        total_element = root.find(".//ns:total", self._ns)
        icms_tot_element = None
        if total_element is not None:
            icms_tot_element = total_element.find(".//ns:ICMSTot", self._ns)
        
        if icms_tot_element is not None:
            v_prod = self.formatar_valor(self.obter_valor_tag_pai(icms_tot_element, "vProd"))
            v_bc = self.formatar_valor(self.obter_valor_tag_pai(icms_tot_element, "vBC"))
            v_icms = self.formatar_valor(self.obter_valor_tag_pai(icms_tot_element, "vICMS"))
            v_ipi = self.formatar_valor(self.obter_valor_tag_pai(icms_tot_element, "vIPI"))
            v_bcst = self.formatar_valor(self.obter_valor_tag_pai(icms_tot_element, "vBCST"))
            v_st = self.formatar_valor(self.obter_valor_tag_pai(icms_tot_element, "vST"))
            v_nf = self.formatar_valor(self.obter_valor_tag_pai(icms_tot_element, "vNF"))
        else:
            v_prod = self.formatar_valor(self.obter_valor_tag(root, "vProd"))
            v_bc = self.formatar_valor(self.obter_valor_tag(root, "vBC"))
            v_icms = self.formatar_valor(self.obter_valor_tag(root, "vICMS"))
            v_ipi = self.formatar_valor(self.obter_valor_tag(root, "vIPI"))
            v_bcst = self.formatar_valor(self.obter_valor_tag(root, "vBCST"))
            v_st = self.formatar_valor(self.obter_valor_tag(root, "vST"))
            v_nf = self.formatar_valor(self.obter_valor_tag(root, "vNF"))

        return {
            "nNF": n_nf, "serie": serie, "dhEmi": dh_emi, "chNFe": ch_nfe,
            "cnpj_emit": cnpj_emit, "xNome_emit": x_nome_emit, "ie_emit": ie_emit,
            "cnpj_dest": cnpj_dest, "xNome_dest": x_nome_dest, "xMun_dest": x_mun_dest,
            "vProd": v_prod, "vBC": v_bc, "vICMS": v_icms, "vIPI": v_ipi,
            "vBCST": v_bcst, "vST": v_st, "vNF": v_nf
        }

    def extrair_duplicatas(self, root: ET.Element) -> list[dict[str, Any]]:
        """Extrai todas as duplicatas da Nota Fiscal."""
        duplicatas: list[dict[str, Any]] = []
        for dup in root.findall(".//ns:dup", self._ns):
            n_dup = self.obter_valor_tag(dup, "nDup")
            d_venc = self.formatar_data_br(self.obter_valor_tag(dup, "dVenc"))
            v_dup = self.formatar_valor(self.obter_valor_tag(dup, "vDup"))
            duplicatas.append({"nDup": n_dup, "dVenc": d_venc, "vDup": v_dup})
        return duplicatas

    def extrair_produtos(self, root: ET.Element) -> list[dict[str, Any]]:
        """Extrai todos os produtos da Nota Fiscal."""
        produtos: list[dict[str, Any]] = []
        for det in root.findall(".//ns:det", self._ns):
            prod = det.find("ns:prod", self._ns)
            imposto = det.find("ns:imposto", self._ns)

            if prod is not None:
                item = {
                    "CODIGO_PRODUTO": self.obter_valor_tag(prod, "cProd"),
                    "EAN": self.obter_valor_tag(prod, "cEAN"),
                    "DESCRIÇÃO": self.obter_valor_tag(prod, "xProd"),
                    "NCM": self.obter_valor_tag(prod, "NCM"),
                    "CEST": self.obter_valor_tag(prod, "CEST", ""),
                    "CFOP": self.obter_valor_tag(prod, "CFOP"),
                    "QUANTIDADE": self.formatar_valor(self.obter_valor_tag(prod, "qCom")),
                    "V.UNIT": self.formatar_valor(self.obter_valor_tag(prod, "vUnCom")),
                    "V.TOT": self.formatar_valor(self.obter_valor_tag(prod, "vProd")),
                    "B.ICM": self.formatar_valor(self.obter_valor_tag(imposto, "vBC") if imposto else 0),
                    "V.ICM": self.formatar_valor(self.obter_valor_tag(imposto, "vICMS") if imposto else 0),
                    "AL.ICM": self.formatar_valor(self.obter_valor_tag(imposto, "pICMS") if imposto else 0),
                    "MVA": self.formatar_valor(self.obter_valor_tag(imposto, "pMVAST") if imposto else 0),
                    "B.ST": self.formatar_valor(self.obter_valor_tag(imposto, "vBCST") if imposto else 0),
                    "ICMSTD": self.formatar_valor(self.obter_valor_tag(imposto, "pICMSST") if imposto else 0),
                    "ST": self.formatar_valor(self.obter_valor_tag(imposto, "vICMSST") if imposto else 0),
                    "V.IPI": self.formatar_valor(self.obter_valor_tag(imposto, "vIPI") if imposto else 0),
                    "ALI.IPI": self.formatar_valor(self.obter_valor_tag(imposto, "pIPI") if imposto else 0),
                }
                produtos.append(item)
        return produtos


class ServicoConversaoExcel:
    """Gerencia a criação e formatação dos arquivos Excel gerados."""
    
    def __init__(self, logger: logging.Logger) -> None:
        self._logger = logger

    def _aplicar_formatacao_br(self, worksheet: Any, num_linhas: int, colunas_indice: list[int]) -> None:
        """Aplica formato brasileiro (vírgula decimal) em colunas de uma aba específica."""
        if not openpyxl or not numbers:
            return
        formato_br = numbers.FORMAT_NUMBER_COMMA_SEPARATED1
        for col_idx in colunas_indice:
            letra_col = get_column_letter(col_idx)
            for row in range(2, num_linhas + 2):
                celula = worksheet[f"{letra_col}{row}"]
                celula.number_format = formato_br

    def criar_excel_com_abas(self, df_geral: pd.DataFrame, df_dup: pd.DataFrame, df_prod: pd.DataFrame, caminho_saida: str) -> None:
        """Gera arquivo Excel contendo até 3 abas."""
        with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
            df_geral.to_excel(writer, sheet_name="Dados Gerais", index=False)
            
            # Formatação Dados Gerais
            colunas_geral = [
                "Valor Total Produtos", "Base ICMS", "Valor ICMS",
                "Valor IPI", "Base ICMS ST", "Valor ICMS ST", "Valor Total NF"
            ]
            indices_geral = [df_geral.columns.get_loc(c) + 1 for c in colunas_geral if c in df_geral.columns]
            self._aplicar_formatacao_br(writer.sheets["Dados Gerais"], len(df_geral), indices_geral)

            # Duplicatas
            df_dup.to_excel(writer, sheet_name="Duplicatas", index=False)
            if "Valor Duplicata" in df_dup.columns:
                indices_dup = [df_dup.columns.get_loc("Valor Duplicata") + 1]
                self._aplicar_formatacao_br(writer.sheets["Duplicatas"], len(df_dup), indices_dup)
            elif "vDup" in df_dup.columns:
                indices_dup = [df_dup.columns.get_loc("vDup") + 1]
                self._aplicar_formatacao_br(writer.sheets["Duplicatas"], len(df_dup), indices_dup)

            # Produtos
            df_prod.to_excel(writer, sheet_name="Produtos", index=False)
            colunas_prod = [
                "QUANTIDADE", "V.UNIT", "V.TOT", "B.ICM", "V.ICM",
                "AL.ICM", "MVA", "B.ST", "ICMSTD", "ST", "V.IPI", "ALI.IPI"
            ]
            indices_prod = [df_prod.columns.get_loc(c) + 1 for c in colunas_prod if c in df_prod.columns]
            self._aplicar_formatacao_br(writer.sheets["Produtos"], len(df_prod), indices_prod)


class ProcessadorNFe:
    """Orquestra a conversão de XML para Excel."""
    
    def __init__(self, logger: logging.Logger) -> None:
        self._logger = logger
        self._extrator = ServicoExtracaoNFe(logger)
        self._excel = ServicoConversaoExcel(logger)

    def processar_individual(self, caminho_xml: str, diretorio_saida: str) -> str:
        """Lê um XML e gera um Excel correspondente. Retorna o nome do arquivo gerado."""
        try:
            tree = ET.parse(caminho_xml)
            root = tree.getroot()

            nf_data = self._extrator.extrair_dados_gerais(root)
            duplicatas = self._extrator.extrair_duplicatas(root)
            produtos = self._extrator.extrair_produtos(root)

            n_nf = nf_data.get("nNF", "SEM_NUMERO")
            nome_emit = nf_data.get("xNome_emit", "SEM_NOME")
            nome_emit_limpo = "".join(c for c in nome_emit if c.isalnum() or c in (' ', '-', '_')).strip()

            nome_arquivo = f"{n_nf}_{nome_emit_limpo}.xlsx"
            caminho_saida = os.path.join(diretorio_saida, nome_arquivo)

            df_geral = pd.DataFrame([{
                "Número NF": nf_data["nNF"], "Série": nf_data["serie"],
                "Data Emissão": nf_data["dhEmi"], "Chave NFe": nf_data["chNFe"],
                "CNPJ Emitente": nf_data["cnpj_emit"], "Nome Emitente": nf_data["xNome_emit"],
                "IE Emitente": nf_data["ie_emit"], "CNPJ Destinatário": nf_data["cnpj_dest"],
                "Nome Destinatário": nf_data["xNome_dest"], "Município Destinatário": nf_data["xMun_dest"],
                "Valor Total Produtos": nf_data["vProd"], "Base ICMS": nf_data["vBC"],
                "Valor ICMS": nf_data["vICMS"], "Valor IPI": nf_data["vIPI"],
                "Base ICMS ST": nf_data["vBCST"], "Valor ICMS ST": nf_data["vST"],
                "Valor Total NF": nf_data["vNF"]
            }])

            df_dup = pd.DataFrame(duplicatas) if duplicatas else pd.DataFrame({"Mensagem": ["Sem duplicatas"]})
            df_prod = pd.DataFrame(produtos) if produtos else pd.DataFrame({"Mensagem": ["Sem produtos"]})

            self._excel.criar_excel_com_abas(df_geral, df_dup, df_prod, caminho_saida)
            return nome_arquivo

        except Exception as erro:
            self._logger.exception("Falha ao processar arquivo %s", caminho_xml)
            raise ErroConversao(f"Erro ao processar: {erro}") from erro

    def processar_unico(self, arquivos_xml: list[str], dir_xml: str, dir_saida: str, callback_progresso: Callable[[float, str], None]) -> str:
        """Gera um único Excel contendo as informações de múltiplos XMLs."""
        try:
            caminho_saida = os.path.join(dir_saida, "Relatorio_NFe_Completo.xlsx")
            
            todos_geral: list[dict[str, Any]] = []
            todos_dup: list[dict[str, Any]] = []
            todos_prod: list[dict[str, Any]] = []

            for index, arquivo in enumerate(arquivos_xml):
                caminho_xml = os.path.join(dir_xml, arquivo)
                tree = ET.parse(caminho_xml)
                root = tree.getroot()

                nf_data = self._extrator.extrair_dados_gerais(root)
                duplicatas = self._extrator.extrair_duplicatas(root)
                produtos = self._extrator.extrair_produtos(root)

                identificador = f"NF {nf_data['nNF']} - {nf_data['xNome_emit']}"

                geral_row = {
                    "Identificador": identificador, "Número NF": nf_data["nNF"], "Série": nf_data["serie"],
                    "Data Emissão": nf_data["dhEmi"], "Chave NFe": nf_data["chNFe"],
                    "CNPJ Emitente": nf_data["cnpj_emit"], "Nome Emitente": nf_data["xNome_emit"],
                    "IE Emitente": nf_data["ie_emit"], "CNPJ Destinatário": nf_data["cnpj_dest"],
                    "Nome Destinatário": nf_data["xNome_dest"], "Município Destinatário": nf_data["xMun_dest"],
                    "Valor Total Produtos": nf_data["vProd"], "Base ICMS": nf_data["vBC"],
                    "Valor ICMS": nf_data["vICMS"], "Valor IPI": nf_data["vIPI"],
                    "Base ICMS ST": nf_data["vBCST"], "Valor ICMS ST": nf_data["vST"],
                    "Valor Total NF": nf_data["vNF"]
                }
                todos_geral.append(geral_row)

                for dup in duplicatas:
                    dup_row = {
                        "Identificador": identificador, "Número NF": nf_data["nNF"],
                        "Nome Emitente": nf_data["xNome_emit"], "Número Duplicata": dup["nDup"],
                        "Vencimento": dup["dVenc"], "Valor Duplicata": dup["vDup"]
                    }
                    todos_dup.append(dup_row)

                for prod in produtos:
                    prod_row = {
                        "Identificador": identificador, "Número NF": nf_data["nNF"],
                        "Nome Emitente": nf_data["xNome_emit"]
                    }
                    prod_row.update(prod)
                    todos_prod.append(prod_row)

                # Atualiza interface
                progresso = (index + 1) / len(arquivos_xml)
                callback_progresso(progresso, f"Processando: {index + 1}/{len(arquivos_xml)} - {arquivo}")

            df_geral = pd.DataFrame(todos_geral)
            df_dup = pd.DataFrame(todos_dup) if todos_dup else pd.DataFrame({"Mensagem": ["Sem duplicatas"]})
            df_prod = pd.DataFrame(todos_prod) if todos_prod else pd.DataFrame({"Mensagem": ["Sem produtos"]})

            self._excel.criar_excel_com_abas(df_geral, df_dup, df_prod, caminho_saida)
            return "Relatorio_NFe_Completo.xlsx"

        except Exception as erro:
            self._logger.exception("Falha ao processar relatório único")
            raise ErroConversao(f"Erro ao processar relatório único: {erro}") from erro


# =============================================================================
# INTERFACE (UI)
# =============================================================================

class InterfaceConversor(ctk.CTk):
    """Janela principal da aplicação, utilizando CustomTkinter de forma isolada."""

    def __init__(self) -> None:
        super().__init__()
        
        # Configuração Básica
        self.title("Conversor NFe XML para Excel")
        self.geometry("700x550")
        self.resizable(False, False)
        self.configure(fg_color=FUNDO_PRINCIPAL)

        # Threading e Comunicação
        self._fila_ui: queue.Queue[dict[str, Any]] = queue.Queue()
        self._processador = ProcessadorNFe(logger)
        self._processando = False

        # Variáveis de Estado
        self.var_xml_dir = ctk.StringVar()
        self.var_export_dir = ctk.StringVar()
        self.var_processo_opt = ctk.StringVar(value="individual")

        self._configurar_grid()
        self._construir_interface()
        self._iniciar_loop_fila()

    def _configurar_grid(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1) # Conteudo
        self.grid_rowconfigure(1, weight=0) # Rodapé

    def _construir_interface(self) -> None:
        # Frame Principal Elevado
        frame_main = ctk.CTkFrame(self, fg_color=SUPERFICIE, corner_radius=RAIO_BORDA)
        frame_main.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
        frame_main.grid_columnconfigure(0, weight=1)

        # Cabeçalho
        lbl_titulo = ctk.CTkLabel(
            frame_main, text="Extrator de Dados NFe (Mod 55)",
            font=ctk.CTkFont(size=22, weight="bold"), text_color=OURO_PRINCIPAL
        )
        lbl_titulo.grid(row=0, column=0, pady=(20, 10))

        lbl_desc = ctk.CTkLabel(
            frame_main, text="Selecione os diretórios e o modo de conversão.",
            font=ctk.CTkFont(size=14), text_color=TEXTO_SECUNDARIO
        )
        lbl_desc.grid(row=1, column=0, pady=(0, 20))

        # Configurações de Diretório
        frame_dir = ctk.CTkFrame(frame_main, fg_color="transparent")
        frame_dir.grid(row=2, column=0, sticky="ew", padx=20, pady=10)
        frame_dir.grid_columnconfigure(1, weight=1)

        # XML
        btn_xml = ctk.CTkButton(
            frame_dir, text="Pasta XML", width=120,
            fg_color=BORDA_FORTE, hover_color=BORDA_SUTIL, text_color=TEXTO_DESTAQUE,
            command=self._selecionar_pasta_xml
        )
        btn_xml.grid(row=0, column=0, padx=(0, 10), pady=5)
        
        entry_xml = ctk.CTkEntry(
            frame_dir, textvariable=self.var_xml_dir,
            fg_color=FUNDO_PRINCIPAL, border_color=BORDA_FORTE, text_color=TEXTO_PRIMARIO
        )
        entry_xml.grid(row=0, column=1, sticky="ew", pady=5)
        entry_xml.configure(state="readonly")

        # Destino
        btn_destino = ctk.CTkButton(
            frame_dir, text="Destino", width=120,
            fg_color=BORDA_FORTE, hover_color=BORDA_SUTIL, text_color=TEXTO_DESTAQUE,
            command=self._selecionar_pasta_destino
        )
        btn_destino.grid(row=1, column=0, padx=(0, 10), pady=5)

        entry_destino = ctk.CTkEntry(
            frame_dir, textvariable=self.var_export_dir,
            fg_color=FUNDO_PRINCIPAL, border_color=BORDA_FORTE, text_color=TEXTO_PRIMARIO
        )
        entry_destino.grid(row=1, column=1, sticky="ew", pady=5)
        entry_destino.configure(state="readonly")

        # Opções de Processamento
        frame_opt = ctk.CTkFrame(frame_main, fg_color=FUNDO_PRINCIPAL, corner_radius=RAIO_BORDA)
        frame_opt.grid(row=3, column=0, sticky="ew", padx=20, pady=15)
        
        lbl_opt = ctk.CTkLabel(frame_opt, text="Modo de Saída:", font=ctk.CTkFont(weight="bold"), text_color=TEXTO_DESTAQUE)
        lbl_opt.pack(pady=(10, 5))

        rb_indiv = ctk.CTkRadioButton(
            frame_opt, text="Um Excel por XML", variable=self.var_processo_opt, value="individual",
            text_color=TEXTO_PRIMARIO, fg_color=ESMERALDA_PRIMARIA, hover_color=ESMERALDA_DEEP
        )
        rb_indiv.pack(side="left", expand=True, pady=(0, 10))

        rb_unico = ctk.CTkRadioButton(
            frame_opt, text="Único Excel Combinado", variable=self.var_processo_opt, value="unico",
            text_color=TEXTO_PRIMARIO, fg_color=ESMERALDA_PRIMARIA, hover_color=ESMERALDA_DEEP
        )
        rb_unico.pack(side="right", expand=True, pady=(0, 10))

        # Controles
        self.barra_progresso = ctk.CTkProgressBar(
            frame_main, progress_color=ESMERALDA_PRIMARIA, fg_color=BORDA_FORTE, height=8
        )
        self.barra_progresso.grid(row=4, column=0, sticky="ew", padx=20, pady=(20, 10))
        self.barra_progresso.set(0)

        self.lbl_status = ctk.CTkLabel(
            frame_main, text="Aguardando configuração...",
            font=ctk.CTkFont(size=12), text_color=TEXTO_SECUNDARIO
        )
        self.lbl_status.grid(row=5, column=0, pady=(0, 10))

        self.btn_iniciar = ctk.CTkButton(
            frame_main, text="Iniciar Conversão", height=40,
            fg_color=ESMERALDA_PRIMARIA, hover_color=ESMERALDA_DEEP,
            font=ctk.CTkFont(weight="bold", size=14),
            command=self._iniciar_processamento
        )
        self.btn_iniciar.grid(row=6, column=0, pady=(10, 20), padx=40, sticky="ew")

        # Rodapé
        rodape = ctk.CTkLabel(
            self, text="Roberto Santos [LABS]©", font=ctk.CTkFont(size=10), text_color=TEXTO_SECUNDARIO
        )
        rodape.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 8))

    def _selecionar_pasta_xml(self) -> None:
        caminho = filedialog.askdirectory(title="Selecione a pasta com arquivos XML")
        if caminho:
            self.var_xml_dir.set(caminho)

    def _selecionar_pasta_destino(self) -> None:
        caminho = filedialog.askdirectory(title="Selecione o destino para os Excel")
        if caminho:
            self.var_export_dir.set(caminho)

    def _iniciar_loop_fila(self) -> None:
        """Processa as mensagens vindas das threads em background."""
        try:
            while True:
                msg = self._fila_ui.get_nowait()
                self._tratar_mensagem(msg)
        except queue.Empty:
            pass
        self.after(50, self._iniciar_loop_fila)

    def _tratar_mensagem(self, msg: dict[str, Any]) -> None:
        tipo = msg.get("tipo")
        
        if tipo == "progresso":
            valor = msg.get("valor", 0.0)
            texto = msg.get("texto", "")
            self.barra_progresso.set(valor)
            self.lbl_status.configure(text=texto, text_color=TEXTO_SECUNDARIO)
            
        elif tipo == "sucesso":
            texto = msg.get("texto", "Concluído")
            self.lbl_status.configure(text=texto, text_color=ESMERALDA_SUCESSO)
            self._restaurar_estado_inicial()
            messagebox.showinfo("Sucesso", msg.get("detalhe", texto))
            
        elif tipo == "erro":
            texto = msg.get("texto", "Falha")
            self.lbl_status.configure(text=texto, text_color=ERRO)
            self.barra_progresso.set(0)
            self._restaurar_estado_inicial()
            messagebox.showerror("Erro", msg.get("detalhe", texto))

    def _restaurar_estado_inicial(self) -> None:
        self._processando = False
        self.btn_iniciar.configure(state="normal", text="Iniciar Conversão")
        self.barra_progresso.configure(mode="determinate")

    def _iniciar_processamento(self) -> None:
        if self._processando:
            return
            
        dir_xml = self.var_xml_dir.get()
        dir_destino = self.var_export_dir.get()
        modo = self.var_processo_opt.get()

        if not dir_xml or not dir_destino:
            self._fila_ui.put({"tipo": "erro", "texto": "Selecione as pastas primeiro.", "detalhe": "Diretório de origem ou destino não selecionado."})
            return

        try:
            arquivos = [f for f in os.listdir(dir_xml) if f.lower().endswith(".xml")]
        except Exception as e:
            self._fila_ui.put({"tipo": "erro", "texto": "Erro ao ler diretório.", "detalhe": str(e)})
            return

        if not arquivos:
            self._fila_ui.put({"tipo": "erro", "texto": "Nenhum XML encontrado.", "detalhe": "A pasta selecionada não contém arquivos XML."})
            return

        self._processando = True
        self.btn_iniciar.configure(state="disabled", text="Processando...")
        self.barra_progresso.set(0)

        threading.Thread(
            target=self._executar_conversao,
            args=(arquivos, dir_xml, dir_destino, modo),
            daemon=True
        ).start()

    def _executar_conversao(self, arquivos: list[str], dir_xml: str, dir_destino: str, modo: str) -> None:
        logger.info("Iniciando conversão de %d arquivos no modo %s", len(arquivos), modo)
        
        sucessos = 0
        erros_log = []

        try:
            if modo == "individual":
                for i, arquivo in enumerate(arquivos):
                    caminho_completo = os.path.join(dir_xml, arquivo)
                    self._fila_ui.put({
                        "tipo": "progresso",
                        "valor": (i + 1) / len(arquivos),
                        "texto": f"Processando: {i + 1}/{len(arquivos)} - {arquivo}"
                    })
                    
                    try:
                        self._processador.processar_individual(caminho_completo, dir_destino)
                        sucessos += 1
                    except Exception as e:
                        erros_log.append(f"{arquivo}: {e}")

                if erros_log:
                    detalhes = f"Sucesso: {sucessos} | Erros: {len(erros_log)}\n" + "\n".join(erros_log[:5])
                    tipo = "erro" if sucessos == 0 else "sucesso"
                    self._fila_ui.put({"tipo": tipo, "texto": "Processamento com ressalvas.", "detalhe": detalhes})
                else:
                    self._fila_ui.put({"tipo": "sucesso", "texto": "Todos os arquivos processados!", "detalhe": f"{sucessos} arquivos convertidos com sucesso."})

            elif modo == "unico":
                def callback_progresso(valor: float, texto: str) -> None:
                    self._fila_ui.put({"tipo": "progresso", "valor": valor, "texto": texto})
                
                self._fila_ui.put({"tipo": "progresso", "valor": 0, "texto": "Preparando relatório único..."})
                nome_saida = self._processador.processar_unico(arquivos, dir_xml, dir_destino, callback_progresso)
                self._fila_ui.put({"tipo": "sucesso", "texto": "Relatório unificado gerado!", "detalhe": f"Arquivo {nome_saida} salvo com sucesso."})

        except Exception as e:
            logger.exception("Falha fatal na thread de processamento.")
            self._fila_ui.put({"tipo": "erro", "texto": "Erro crítico no processo.", "detalhe": str(e)})


# =============================================================================
# INICIALIZAÇÃO
# =============================================================================

def main() -> None:
    # Configuração global CustomTkinter
    ctk.set_appearance_mode("system")
    ctk.set_default_color_theme("blue")
    
    logger.info("Iniciando aplicação NFe to Excel.")
    try:
        app = InterfaceConversor()
        app.mainloop()
    except Exception:
        logger.exception("Erro não tratado que causou o fechamento da aplicação.")

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
    logger.info("Iniciando instância da sub-aplicação NFe Excel")
    try:
        main()
        logger.info("Sub-aplicação NFe Excel encerrada normalmente")
    except Exception:
        logger.exception("Falha crítica na execução da sub-aplicação NFe Excel")
        sys.exit(1)
