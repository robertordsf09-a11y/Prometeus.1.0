from __future__ import annotations

import logging
import os
import queue
import re
import sys
import threading
from datetime import datetime
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

try:
    import customtkinter as ctk
    from tkinter import filedialog, messagebox
except ImportError as erro:
    print(f"Erro de importação GUI: {erro}")
    print("Instale: pip install customtkinter pillow")
    sys.exit(1)

# Fiscal
try:
    from brazilfiscalreport.danfe import Danfe, DanfeConfig, ReceiptPosition
    from brazilfiscalreport.dacte import Dacte, DacteConfig
    FISCAL_OK: bool = True
except ImportError:
    FISCAL_OK = False

# Excel
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    EXCEL_OK: bool = True
except ImportError:
    EXCEL_OK = False

def obter_diretorio_base() -> str:
    """
    Retorna o diretório raiz da aplicação.
    Compatível com execução direta (.py) e executável compilado via Nuitka.
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE_DIR: str = obter_diretorio_base()

def criar_logger(nome_modulo: str, usuario: str = "sistema") -> logging.Logger:
    """Cria logger configurado com formato padrão e rotação de arquivo."""
    formato = f"[%(asctime)s],[{usuario}],[{nome_modulo}] %(levelname)s: %(message)s"
    formatador = logging.Formatter(formato, datefmt="%Y-%m-%d %H:%M:%S")

    caminho_log = os.path.join(BASE_DIR, "logs", "conversor_xmls.log")
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

logger: logging.Logger = criar_logger("conversor_xmls")

# Tokens Visuais Premium
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
AVISO: str = "#FFB800"

NS: dict[str, str] = {
    "nfe": "http://www.portalfiscal.inf.br/nfe",
    "cte": "http://www.portalfiscal.inf.br/cte",
}

class ErroFiscal(Exception):
    """Exceção levantada quando houver erro no processamento fiscal (XML/PDF/Excel)."""
    pass

def _ns(tag: str, ns: str) -> str:
    """Formata tag XML com namespace."""
    return f"{{{NS[ns]}}}{tag}"

def detectar_tipo(caminho: Path) -> str:
    """Detecta o tipo de XML (nfe, cte ou desconhecido)."""
    try:
        tree = ET.parse(caminho)
        root = tree.getroot()
        raw = root.tag
        ns = raw.split("}")[0].lstrip("{") if "}" in raw else ""
        if NS["nfe"] in ns:
            return "nfe"
        if NS["cte"] in ns:
            return "cte"
        for el in root.iter():
            t = el.tag.split("}")[-1] if "}" in el.tag else el.tag
            if t == "infNFe":
                return "nfe"
            if t in ("infCte", "infCTe"):
                return "cte"
    except Exception as erro:
        logger.exception("detectar_tipo | falha ao parsear XML | caminho=%s | erro=%s", caminho, erro)
    return "desconhecido"

def _find(root: ET.Element, *caminhos: str, ns_key: str = "nfe") -> str:
    """Busca um valor em múltiplos caminhos possíveis dentro da árvore XML."""
    for path in caminhos:
        partes = path.split("/")
        el: ET.Element | None = root
        for p in partes:
            if el is None:
                break
            el = el.find(_ns(p, ns_key))
        if el is not None and el.text:
            return el.text.strip()
    return ""

def formatar_cnpj_cpf(v: str) -> str:
    """Formata string numérica como CNPJ ou CPF."""
    v = re.sub(r"\D", "", v)
    if len(v) == 14:
        return f"{v[:2]}.{v[2:5]}.{v[5:8]}/{v[8:12]}-{v[12:]}"
    if len(v) == 11:
        return f"{v[:3]}.{v[3:6]}.{v[6:9]}-{v[9:]}"
    return v

def formatar_data(v: str) -> str:
    """Formata data ISO para dd/mm/yyyy."""
    if not v:
        return ""
    try:
        return datetime.fromisoformat(v[:10]).strftime("%d/%m/%Y")
    except ValueError:
        return v[:10]

def extrair_nfe(caminho: Path) -> dict[str, Any]:
    """Extrai informações relevantes de um XML de NF-e."""
    try:
        tree = ET.parse(caminho)
        root = tree.getroot()

        def find_infnfe(el: ET.Element) -> ET.Element | None:
            t = el.tag.split("}")[-1] if "}" in el.tag else el.tag
            if t == "infNFe":
                return el
            for child in el:
                r = find_infnfe(child)
                if r is not None:
                    return r
            return None

        inf = find_infnfe(root)
        if inf is None:
            return {}

        def f(*paths: str) -> str:
            return _find(inf, *paths, ns_key="nfe")

        chave = inf.get("Id", "").replace("NFe", "")
        produtos: list[dict[str, Any]] = []

        for det in inf.findall(_ns("det", "nfe")):
            prod = det.find(_ns("prod", "nfe"))
            if prod is None:
                continue
            imposto = det.find(_ns("imposto", "nfe"))
            icms_val = ""
            ipi_val = ""
            if imposto:
                for icms_grp in imposto.findall(_ns("ICMS", "nfe")):
                    for child in icms_grp:
                        vv = child.find(_ns("vICMS", "nfe"))
                        if vv is not None and vv.text:
                            icms_val = vv.text
                ipi_grp = imposto.find(_ns("IPI", "nfe"))
                if ipi_grp is not None:
                    for child in ipi_grp:
                        vv = child.find(_ns("vIPI", "nfe"))
                        if vv is not None and vv.text:
                            ipi_val = vv.text

            def fp(tag: str) -> str:
                el_prod = prod.find(_ns(tag, "nfe"))
                return el_prod.text.strip() if el_prod is not None and el_prod.text else ""

            produtos.append({
                "item": det.get("nItem", ""),
                "codigo": fp("cProd"),
                "descricao": fp("xProd"),
                "ncm": fp("NCM"),
                "cfop": fp("CFOP"),
                "unidade": fp("uCom"),
                "quantidade": fp("qCom"),
                "vl_unitario": fp("vUnCom"),
                "vl_total": fp("vProd"),
                "icms": icms_val,
                "ipi": ipi_val,
            })

        dups: list[dict[str, Any]] = []
        cobr = inf.find(_ns("cobr", "nfe"))
        if cobr:
            for dup in cobr.findall(_ns("dup", "nfe")):
                def fd(tag: str) -> str:
                    el_dup = dup.find(_ns(tag, "nfe"))
                    return el_dup.text.strip() if el_dup is not None and el_dup.text else ""
                dups.append({
                    "numero": fd("nDup"),
                    "venc": formatar_data(fd("dVenc")),
                    "valor": fd("vDup"),
                })

        return {
            "tipo": "NF-e",
            "chave": chave,
            "numero": f("ide/nNF"),
            "serie": f("ide/serie"),
            "data_emissao": formatar_data(f("ide/dhEmi", "ide/dEmi")),
            "nat_operacao": f("ide/natOp"),
            "tipo_nf": "Saída" if f("ide/tpNF") == "1" else "Entrada",
            "emit_cnpj": formatar_cnpj_cpf(f("emit/CNPJ", "emit/CPF")),
            "emit_nome": f("emit/xNome"),
            "emit_ie": f("emit/IE"),
            "emit_uf": f("emit/enderEmit/UF"),
            "dest_cnpj": formatar_cnpj_cpf(f("dest/CNPJ", "dest/CPF")),
            "dest_nome": f("dest/xNome"),
            "dest_ie": f("dest/IE"),
            "dest_uf": f("dest/enderDest/UF"),
            "vl_produtos": f("total/ICMSTot/vProd"),
            "vl_frete": f("total/ICMSTot/vFrete"),
            "vl_desconto": f("total/ICMSTot/vDesc"),
            "vl_icms": f("total/ICMSTot/vICMS"),
            "vl_ipi": f("total/ICMSTot/vIPI"),
            "vl_pis": f("total/ICMSTot/vPIS"),
            "vl_cofins": f("total/ICMSTot/vCOFINS"),
            "vl_total_nf": f("total/ICMSTot/vNF"),
            "transp_modalidade": f("transp/modFrete"),
            "transp_nome": f("transp/transporta/xNome"),
            "inf_adicionais": f("infAdic/infCpl"),
            "produtos": produtos,
            "duplicatas": dups,
            "arquivo": caminho.name,
        }
    except Exception as erro:
        logger.exception("extrair_nfe | falha ao processar xml | caminho=%s", caminho)
        raise ErroFiscal(f"Falha ao extrair NF-e do arquivo {caminho.name}") from erro

def extrair_cte(caminho: Path) -> dict[str, Any]:
    """Extrai informações relevantes de um XML de CT-e."""
    try:
        tree = ET.parse(caminho)
        root = tree.getroot()

        def find_infcte(el: ET.Element) -> ET.Element | None:
            t = el.tag.split("}")[-1] if "}" in el.tag else el.tag
            if t in ("infCte", "infCTe"):
                return el
            for child in el:
                r = find_infcte(child)
                if r is not None:
                    return r
            return None

        inf = find_infcte(root)
        if inf is None:
            return {}

        def f(*paths: str) -> str:
            return _find(inf, *paths, ns_key="cte")

        chave = inf.get("Id", "").replace("CTe", "")

        return {
            "tipo": "CT-e",
            "chave": chave,
            "numero": f("ide/nCT"),
            "serie": f("ide/serie"),
            "data_emissao": formatar_data(f("ide/dhEmi", "ide/dEmi")),
            "cfop": f("ide/CFOP"),
            "nat_operacao": f("ide/natOp", "ide/descServico"),
            "emit_cnpj": formatar_cnpj_cpf(f("emit/CNPJ")),
            "emit_nome": f("emit/xNome"),
            "emit_ie": f("emit/IE"),
            "emit_uf": f("emit/enderEmit/UF"),
            "rem_cnpj": formatar_cnpj_cpf(f("rem/CNPJ", "rem/CPF")),
            "rem_nome": f("rem/xNome"),
            "dest_cnpj": formatar_cnpj_cpf(f("dest/CNPJ", "dest/CPF")),
            "dest_nome": f("dest/xNome"),
            "dest_uf": f("dest/enderDest/UF"),
            "tom_cnpj": formatar_cnpj_cpf(f("toma3/CNPJ", "toma4/CNPJ", "toma3/CPF")),
            "tom_nome": f("toma3/xNome", "toma4/xNome"),
            "vl_total": f("vPrest/vTPrest"),
            "vl_receber": f("vPrest/vRec"),
            "vl_carga": f("infCTeNorm/infCarga/vCarga", "infCarga/vCarga"),
            "modal": f("ide/modal"),
            "inf_adicionais": f("compl/xObs"),
            "arquivo": caminho.name,
        }
    except Exception as erro:
        logger.exception("extrair_cte | falha ao processar xml | caminho=%s", caminho)
        raise ErroFiscal(f"Falha ao extrair CT-e do arquivo {caminho.name}") from erro

def _estilo_cabecalho(ws: Any, row: int, cols: int, titulo: list[str], fill_color: str = "1A6B4A") -> None:
    """Aplica formatação padronizada em cabeçalhos de planilhas."""
    fill = PatternFill("solid", fgColor=fill_color)
    fonte = Font(bold=True, color="FFFFFF", size=10)
    borda = Border(bottom=Side(style="medium", color="FFFFFF"))
    for col, label in zip(range(1, cols + 1), titulo):
        cell = ws.cell(row=row, column=col, value=label)
        cell.font = fonte
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = borda

def _estilo_linha(ws: Any, row: int, valores: list[Any], zebra: bool = False) -> None:
    """Aplica formatação padronizada em linhas de planilhas."""
    fill_color = "F0FFF8" if zebra else "FFFFFF"
    fill = PatternFill("solid", fgColor=fill_color)
    fonte = Font(size=9)
    borda_lado = Side(style="thin", color="D0D0D0")
    borda = Border(left=borda_lado, right=borda_lado, top=borda_lado, bottom=borda_lado)
    for col, val in enumerate(valores, 1):
        cell = ws.cell(row=row, column=col, value=val)
        cell.font = fonte
        cell.fill = fill
        cell.border = borda
        cell.alignment = Alignment(vertical="center")

def _auto_largura(ws: Any, min_w: int = 8, max_w: int = 50) -> None:
    """Ajusta automaticamente a largura das colunas do Excel."""
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                max_len = max(max_len, len(str(cell.value or "")))
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = max(min_w, min(max_len + 2, max_w))

def gerar_excel(registros_nfe: list[dict[str, Any]], registros_cte: list[dict[str, Any]], caminho_saida: Path) -> None:
    """Gera uma planilha Excel com as informações extraídas dos XMLs."""
    try:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        cor_nfe = "0D5C36"
        cor_cte = "0D3A6B"

        if registros_nfe:
            ws_nfe = wb.create_sheet("NF-e — Resumo")
            ws_nfe.row_dimensions[1].height = 30
            cab_nfe = [
                "Arquivo", "Chave NF-e", "Número", "Série", "Data Emissão",
                "Tipo", "Natureza Operação",
                "Emitente CNPJ", "Emitente Nome", "Emitente UF",
                "Destinatário CNPJ", "Destinatário Nome", "Destinatário UF",
                "Vl. Produtos", "Vl. Frete", "Vl. Desconto",
                "Vl. ICMS", "Vl. IPI", "Vl. PIS", "Vl. COFINS",
                "Vl. Total NF",
            ]
            _estilo_cabecalho(ws_nfe, 1, len(cab_nfe), cab_nfe, cor_nfe)

            for i, r in enumerate(registros_nfe, 2):
                def n(k: str) -> float | str:
                    try:
                        v = r.get(k, "")
                        return float(v) if v else 0.0
                    except ValueError:
                        return r.get(k, "")

                linha = [
                    r.get("arquivo", ""), r.get("chave", ""),
                    r.get("numero", ""), r.get("serie", ""), r.get("data_emissao", ""),
                    r.get("tipo_nf", ""), r.get("nat_operacao", ""),
                    r.get("emit_cnpj", ""), r.get("emit_nome", ""), r.get("emit_uf", ""),
                    r.get("dest_cnpj", ""), r.get("dest_nome", ""), r.get("dest_uf", ""),
                    n("vl_produtos"), n("vl_frete"), n("vl_desconto"),
                    n("vl_icms"), n("vl_ipi"), n("vl_pis"), n("vl_cofins"),
                    n("vl_total_nf"),
                ]
                _estilo_linha(ws_nfe, i, linha, i % 2 == 0)

            for row_ in ws_nfe.iter_rows(min_row=2, min_col=14, max_col=21):
                for cell in row_:
                    cell.number_format = 'R$ #,##0.00'
                    cell.alignment = Alignment(horizontal="right", vertical="center")
            _auto_largura(ws_nfe)

            ws_prod = wb.create_sheet("NF-e — Produtos")
            ws_prod.row_dimensions[1].height = 30
            cab_prod = [
                "Arquivo", "NF Número", "Item", "Código", "Descrição",
                "NCM", "CFOP", "Unidade", "Quantidade",
                "Vl. Unitário", "Vl. Total", "ICMS", "IPI",
            ]
            _estilo_cabecalho(ws_prod, 1, len(cab_prod), cab_prod, cor_nfe)
            row_p = 2
            for r in registros_nfe:
                for p in r.get("produtos", []):
                    def pn(k: str) -> float | str:
                        try:
                            v = p.get(k, "")
                            return float(v) if v else 0.0
                        except ValueError:
                            return p.get(k, "")
                    linha_prod = [
                        r.get("arquivo", ""), r.get("numero", ""),
                        p.get("item", ""), p.get("codigo", ""), p.get("descricao", ""),
                        p.get("ncm", ""), p.get("cfop", ""), p.get("unidade", ""),
                        pn("quantidade"), pn("vl_unitario"), pn("vl_total"),
                        pn("icms"), pn("ipi"),
                    ]
                    _estilo_linha(ws_prod, row_p, linha_prod, row_p % 2 == 0)
                    row_p += 1

            for row_ in ws_prod.iter_rows(min_row=2, min_col=10, max_col=13):
                for cell in row_:
                    cell.number_format = 'R$ #,##0.00'
                    cell.alignment = Alignment(horizontal="right", vertical="center")
            _auto_largura(ws_prod)

            ws_dup = wb.create_sheet("NF-e — Duplicatas")
            ws_dup.row_dimensions[1].height = 30
            cab_dup = ["Arquivo", "NF Número", "Nº Duplicata", "Vencimento", "Valor"]
            _estilo_cabecalho(ws_dup, 1, len(cab_dup), cab_dup, cor_nfe)
            row_d = 2
            for r in registros_nfe:
                for d in r.get("duplicatas", []):
                    try:
                        val: float | str = float(d.get("valor", 0))
                    except ValueError:
                        val = d.get("valor", "")
                    linha_dup = [
                        r.get("arquivo", ""), r.get("numero", ""),
                        d.get("numero", ""), d.get("venc", ""), val,
                    ]
                    _estilo_linha(ws_dup, row_d, linha_dup, row_d % 2 == 0)
                    row_d += 1

            for row_ in ws_dup.iter_rows(min_row=2, min_col=5, max_col=5):
                for cell in row_:
                    cell.number_format = 'R$ #,##0.00'
                    cell.alignment = Alignment(horizontal="right", vertical="center")
            _auto_largura(ws_dup)

        if registros_cte:
            ws_cte = wb.create_sheet("CT-e — Resumo")
            ws_cte.row_dimensions[1].height = 30
            cab_cte = [
                "Arquivo", "Chave CT-e", "Número", "Série", "Data Emissão",
                "CFOP", "Modal",
                "Emitente CNPJ", "Emitente Nome", "Emitente UF",
                "Remetente CNPJ", "Remetente Nome",
                "Destinatário CNPJ", "Destinatário Nome", "Destinatário UF",
                "Tomador CNPJ", "Tomador Nome",
                "Vl. Total Prestação", "Vl. a Receber", "Vl. Carga",
            ]
            _estilo_cabecalho(ws_cte, 1, len(cab_cte), cab_cte, cor_cte)

            for i, r in enumerate(registros_cte, 2):
                def cn(k: str) -> float | str:
                    try:
                        v = r.get(k, "")
                        return float(v) if v else 0.0
                    except ValueError:
                        return r.get(k, "")

                linha_cte = [
                    r.get("arquivo", ""), r.get("chave", ""),
                    r.get("numero", ""), r.get("serie", ""), r.get("data_emissao", ""),
                    r.get("cfop", ""), r.get("modal", ""),
                    r.get("emit_cnpj", ""), r.get("emit_nome", ""), r.get("emit_uf", ""),
                    r.get("rem_cnpj", ""), r.get("rem_nome", ""),
                    r.get("dest_cnpj", ""), r.get("dest_nome", ""), r.get("dest_uf", ""),
                    r.get("tom_cnpj", ""), r.get("tom_nome", ""),
                    cn("vl_total"), cn("vl_receber"), cn("vl_carga"),
                ]
                _estilo_linha(ws_cte, i, linha_cte, i % 2 == 0)

            for row_ in ws_cte.iter_rows(min_row=2, min_col=18, max_col=20):
                for cell in row_:
                    cell.number_format = 'R$ #,##0.00'
                    cell.alignment = Alignment(horizontal="right", vertical="center")
            _auto_largura(ws_cte)

        wb.save(caminho_saida)
    except Exception as erro:
        logger.exception("gerar_excel | erro ao salvar planilha | caminho=%s", caminho_saida)
        raise ErroFiscal("Falha ao gerar planilha Excel.") from erro

def converter_pdf(caminho_xml: Path, caminho_pdf: Path, tipo: str, danfe_cfg: Any, dacte_cfg: Any) -> bool:
    """Converte o arquivo XML fiscal em documento PDF (DANFE/DACTE)."""
    if not FISCAL_OK:
        return False
    try:
        with open(caminho_xml, "rb") as arquivo:
            xml_bytes = arquivo.read()

        doc: Any
        if tipo == "nfe":
            doc = Danfe(xml=xml_bytes, config=danfe_cfg)
        else:
            doc = Dacte(xml=xml_bytes, config=dacte_cfg)

        pdf_bytes = doc.output(dest="S")
        if isinstance(pdf_bytes, str):
            pdf_bytes = pdf_bytes.encode("latin-1")
        
        with open(caminho_pdf, "wb") as arquivo_pdf:
            arquivo_pdf.write(pdf_bytes)
        return True
    except Exception as erro:
        logger.exception("converter_pdf | falha na conversão de %s | caminho=%s", tipo, caminho_xml)
        raise ErroFiscal(f"Falha ao gerar PDF de {caminho_xml.name}") from erro

# ── INTERFACE GRÁFICA ─────────────────────────────────────────────────────────

ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")

class AppConversorXML(ctk.CTk):
    """Aplicação principal para conversão de XMLs NFe e CTe."""
    def __init__(self) -> None:
        super().__init__()

        self.title("Conversor NF-e / CT-e")
        self.geometry("860x740")
        self.minsize(800, 680)
        self.configure(fg_color=FUNDO_PRINCIPAL)

        self._fila_ui: queue.Queue[dict[str, Any]] = queue.Queue()
        
        self.pasta_entrada = ctk.StringVar()
        self.pasta_saida = ctk.StringVar()
        self.gerar_pdf = ctk.BooleanVar(value=True)
        self.gerar_excel = ctk.BooleanVar(value=False)
        self.recibo_pos = ctk.StringVar(value="topo")
        self.sobrescrever = ctk.BooleanVar(value=False)
        self.processando = False

        self._configurar_grid()
        self._construir_interface()
        self._iniciar_loop_fila()

    def _configurar_grid(self) -> None:
        """Configura a estrutura principal de layout da janela."""
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

    def _iniciar_loop_fila(self) -> None:
        """Processa mensagens da fila de UI periodicamente."""
        self._processar_fila()

    def _processar_fila(self) -> None:
        """Verifica a fila por atualizações de UI advindas da thread em background."""
        try:
            while True:
                mensagem = self._fila_ui.get_nowait()
                acao = mensagem.get("acao")
                if acao == "log":
                    self._atualizar_log_ui(mensagem.get("texto", ""))
                elif acao == "progresso":
                    self._atualizar_progresso_ui(mensagem.get("valor", 0.0), mensagem.get("status", ""))
                elif acao == "fim":
                    self._finalizar_ui(mensagem)
        except queue.Empty:
            pass
        self.after(50, self._processar_fila)

    def _construir_interface(self) -> None:
        """Renderiza os componentes visuais principais da aplicação."""
        scroll = ctk.CTkScrollableFrame(self, fg_color=FUNDO_PRINCIPAL, scrollbar_button_color=BORDA_FORTE)
        scroll.grid(row=0, column=0, sticky="nsew")
        scroll.grid_columnconfigure(0, weight=1)

        linha = 0
        cabecalho = ctk.CTkFrame(scroll, fg_color=SUPERFICIE, corner_radius=0)
        cabecalho.grid(row=linha, column=0, sticky="ew", pady=(0, 10))
        cabecalho.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(cabecalho, text="⚡ Conversor NF-e / CT-e", font=ctk.CTkFont(size=22, weight="bold"), text_color=TEXTO_DESTAQUE).grid(row=0, column=0, sticky="w", padx=28, pady=(18, 2))
        ctk.CTkLabel(cabecalho, text="Gere DANFE, DACTE e planilhas a partir de XMLs", font=ctk.CTkFont(size=12), text_color=TEXTO_SECUNDARIO).grid(row=1, column=0, sticky="w", padx=28, pady=(0, 16))
        
        linha += 1
        linha = self._construir_secao_pastas(scroll, linha)
        linha = self._construir_secao_formatos(scroll, linha)
        linha = self._construir_secao_opcoes(scroll, linha)

        self.btn_iniciar = ctk.CTkButton(scroll, text="▶ Iniciar Conversão", font=ctk.CTkFont(size=14, weight="bold"), height=50, fg_color=ESMERALDA_PRIMARIA, hover_color=ESMERALDA_DEEP, corner_radius=10, command=self._iniciar_conversao)
        self.btn_iniciar.grid(row=linha, column=0, sticky="ew", padx=24, pady=(8, 4))
        linha += 1

        self._construir_secao_progresso(scroll, linha)
        linha += 1

        self._construir_secao_log(scroll, linha)
        linha += 1

        rodape = ctk.CTkLabel(scroll, text="Roberto Santos [LABS]©", font=ctk.CTkFont(size=10), text_color=TEXTO_SECUNDARIO)
        rodape.grid(row=linha, column=0, sticky="ew", padx=10, pady=(10, 8))

    def _construir_secao_pastas(self, parent: ctk.CTkFrame, linha: int) -> int:
        card = ctk.CTkFrame(parent, fg_color=SUPERFICIE, corner_radius=12)
        card.grid(row=linha, column=0, sticky="ew", padx=24, pady=8)
        card.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(card, text="📂 Pastas", font=ctk.CTkFont(size=12, weight="bold"), text_color=TEXTO_DESTAQUE).grid(row=0, column=0, columnspan=3, sticky="w", padx=16, pady=(12, 6))

        def add_linha(lbl_text: str, var: ctk.StringVar, r: int, cmd: Any) -> None:
            ctk.CTkLabel(card, text=lbl_text, font=ctk.CTkFont(size=11), text_color=TEXTO_SECUNDARIO, width=110, anchor="w").grid(row=r, column=0, sticky="w", padx=(16, 4), pady=4)
            ctk.CTkEntry(card, textvariable=var, fg_color=FUNDO_PRINCIPAL, border_color=BORDA_FORTE, text_color=TEXTO_PRIMARIO, height=36).grid(row=r, column=1, sticky="ew", padx=4, pady=4)
            ctk.CTkButton(card, text="Procurar", width=90, height=36, fg_color=BORDA_FORTE, hover_color=ESMERALDA_DEEP, font=ctk.CTkFont(size=11), command=cmd).grid(row=r, column=2, padx=(4, 16), pady=4)

        add_linha("Pasta de Entrada:", self.pasta_entrada, 1, self._selecionar_entrada)
        add_linha("Pasta de Saída:", self.pasta_saida, 2, self._selecionar_saida)
        return linha + 1

    def _construir_secao_formatos(self, parent: ctk.CTkFrame, linha: int) -> int:
        card = ctk.CTkFrame(parent, fg_color=SUPERFICIE, corner_radius=12)
        card.grid(row=linha, column=0, sticky="ew", padx=24, pady=8)
        card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(card, text="📄 Formato de Saída", font=ctk.CTkFont(size=12, weight="bold"), text_color=TEXTO_DESTAQUE).grid(row=0, column=0, sticky="w", padx=16, pady=(12, 6))

        f = ctk.CTkFrame(card, fg_color="transparent")
        f.grid(row=1, column=0, sticky="w", padx=16, pady=(0, 12))

        ctk.CTkCheckBox(f, text="PDF (DANFE / DACTE)", variable=self.gerar_pdf, fg_color=ESMERALDA_PRIMARIA, text_color=TEXTO_PRIMARIO).grid(row=0, column=0, padx=(0, 24), pady=4)
        
        estado_excel = "normal" if EXCEL_OK else "disabled"
        ctk.CTkCheckBox(f, text="Excel (.xlsx)", variable=self.gerar_excel, state=estado_excel, fg_color=OURO_PRINCIPAL, text_color=TEXTO_PRIMARIO if EXCEL_OK else TEXTO_SECUNDARIO).grid(row=0, column=1, padx=0, pady=4)
        return linha + 1

    def _construir_secao_opcoes(self, parent: ctk.CTkFrame, linha: int) -> int:
        card = ctk.CTkFrame(parent, fg_color=SUPERFICIE, corner_radius=12)
        card.grid(row=linha, column=0, sticky="ew", padx=24, pady=8)
        card.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(card, text="⚙️ Opções", font=ctk.CTkFont(size=12, weight="bold"), text_color=TEXTO_DESTAQUE).grid(row=0, column=0, columnspan=2, sticky="w", padx=16, pady=(12, 6))
        
        ctk.CTkLabel(card, text="Posição do Recibo:", font=ctk.CTkFont(size=11), text_color=TEXTO_SECUNDARIO).grid(row=1, column=0, sticky="w", padx=16, pady=4)
        seg1 = ctk.CTkSegmentedButton(card, values=["Topo", "Rodapé"], command=lambda v: self.recibo_pos.set("topo" if v == "Topo" else "baixo"), selected_color=ESMERALDA_PRIMARIA, unselected_color=FUNDO_PRINCIPAL)
        seg1.set("Topo")
        seg1.grid(row=1, column=1, sticky="w", padx=16, pady=4)

        ctk.CTkLabel(card, text="Arquivos Existentes:", font=ctk.CTkFont(size=11), text_color=TEXTO_SECUNDARIO).grid(row=2, column=0, sticky="w", padx=16, pady=(4, 12))
        seg2 = ctk.CTkSegmentedButton(card, values=["Pular", "Sobrescrever"], command=lambda v: self.sobrescrever.set(v == "Sobrescrever"), selected_color=ESMERALDA_PRIMARIA, unselected_color=FUNDO_PRINCIPAL)
        seg2.set("Pular")
        seg2.grid(row=2, column=1, sticky="w", padx=16, pady=(4, 12))
        return linha + 1

    def _construir_secao_progresso(self, parent: ctk.CTkFrame, linha: int) -> None:
        card = ctk.CTkFrame(parent, fg_color=SUPERFICIE, corner_radius=12)
        card.grid(row=linha, column=0, sticky="ew", padx=24, pady=8)
        card.grid_columnconfigure(0, weight=1)

        self.lbl_status = ctk.CTkLabel(card, text="Aguardando...", font=ctk.CTkFont(size=11), text_color=TEXTO_SECUNDARIO)
        self.lbl_status.grid(row=0, column=0, sticky="w", padx=16, pady=(12, 4))
        
        self.lbl_pct = ctk.CTkLabel(card, text="0%", font=ctk.CTkFont(size=11, weight="bold"), text_color=ESMERALDA_PRIMARIA)
        self.lbl_pct.grid(row=0, column=1, sticky="e", padx=16, pady=(12, 4))
        
        self.barra = ctk.CTkProgressBar(card, height=10, fg_color=FUNDO_PRINCIPAL, progress_color=ESMERALDA_PRIMARIA)
        self.barra.set(0)
        self.barra.grid(row=1, column=0, columnspan=2, sticky="ew", padx=16, pady=(0, 12))

    def _construir_secao_log(self, parent: ctk.CTkFrame, linha: int) -> None:
        card = ctk.CTkFrame(parent, fg_color=SUPERFICIE, corner_radius=12)
        card.grid(row=linha, column=0, sticky="ew", padx=24, pady=8)
        card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(card, text="Log de Processamento", font=ctk.CTkFont(size=11, weight="bold"), text_color=TEXTO_SECUNDARIO).grid(row=0, column=0, sticky="w", padx=16, pady=(10, 4))
        ctk.CTkButton(card, text="Limpar", width=60, height=22, fg_color=BORDA_FORTE, hover_color=FUNDO_PRINCIPAL, font=ctk.CTkFont(size=10), command=self._limpar_log).grid(row=0, column=1, sticky="e", padx=16, pady=(10, 4))

        self.caixa_log = ctk.CTkTextbox(card, height=180, fg_color=FUNDO_PRINCIPAL, text_color=TEXTO_PRIMARIO, font=ctk.CTkFont(family="Courier New", size=10), border_color=BORDA_FORTE, border_width=1)
        self.caixa_log.grid(row=1, column=0, columnspan=2, sticky="ew", padx=16, pady=(0, 12))

    def _selecionar_entrada(self) -> None:
        pasta = filedialog.askdirectory(title="Selecione a pasta com os XMLs")
        if pasta:
            self.pasta_entrada.set(pasta)
            if not self.pasta_saida.get():
                self.pasta_saida.set(pasta)

    def _selecionar_saida(self) -> None:
        pasta = filedialog.askdirectory(title="Selecione a pasta de destino")
        if pasta:
            self.pasta_saida.set(pasta)

    def _limpar_log(self) -> None:
        self.caixa_log.configure(state="normal")
        self.caixa_log.delete("1.0", "end")
        self.caixa_log.configure(state="disabled")

    def _log_thread_safe(self, mensagem: str, nivel: str = "info") -> None:
        if nivel == "error":
            logger.error(mensagem)
        elif nivel == "warning":
            logger.warning(mensagem)
        else:
            logger.info(mensagem)

        ts = datetime.now().strftime("%H:%M:%S")
        texto = f"[{ts}] {mensagem}\n"
        self._fila_ui.put({"acao": "log", "texto": texto})

    def _atualizar_log_ui(self, texto: str) -> None:
        self.caixa_log.configure(state="normal")
        self.caixa_log.insert("end", texto)
        self.caixa_log.see("end")
        self.caixa_log.configure(state="disabled")

    def _progresso_thread_safe(self, valor: float, status: str) -> None:
        self._fila_ui.put({"acao": "progresso", "valor": valor, "status": status})

    def _atualizar_progresso_ui(self, valor: float, status: str) -> None:
        self.barra.set(valor)
        self.lbl_pct.configure(text=f"{int(valor * 100)}%")
        self.lbl_status.configure(text=status)

    def _finalizar_ui(self, dados: dict[str, Any]) -> None:
        self.processando = False
        self.btn_iniciar.configure(text="▶ Iniciar Conversão", state="normal", fg_color=ESMERALDA_PRIMARIA)
        if dados.get("sucesso", False):
            self.lbl_pct.configure(text_color=ESMERALDA_SUCESSO)
            messagebox.showinfo("Concluído", dados.get("msg_final", ""))
        else:
            self.lbl_pct.configure(text_color=ERRO)
            messagebox.showerror("Erro", dados.get("msg_final", ""))

    def _iniciar_conversao(self) -> None:
        if self.processando:
            return

        entrada = self.pasta_entrada.get().strip()
        if not entrada or not Path(entrada).is_dir():
            messagebox.showerror("Erro", "Selecione uma pasta de entrada válida.")
            return
        if not self.gerar_pdf.get() and not self.gerar_excel.get():
            messagebox.showerror("Erro", "Selecione ao menos um formato de saída.")
            return
        if self.gerar_pdf.get() and not FISCAL_OK:
            messagebox.showerror("Erro", "brazilfiscalreport ausente. Instale dependências.")
            return

        self.processando = True
        self.btn_iniciar.configure(text="⏳ Processando...", state="disabled", fg_color=BORDA_FORTE)
        self.lbl_pct.configure(text_color=ESMERALDA_PRIMARIA)
        self._limpar_log()
        
        threading.Thread(target=self._executar_conversao, daemon=True).start()

    def _executar_conversao(self) -> None:
        try:
            entrada = Path(self.pasta_entrada.get().strip())
            saida_str = self.pasta_saida.get().strip()
            saida = Path(saida_str) if saida_str else entrada
            saida.mkdir(parents=True, exist_ok=True)

            logger.info("Iniciando conversão. Entrada: %s | Saída: %s", entrada, saida)

            danfe_cfg = None
            dacte_cfg = None
            if FISCAL_OK and self.gerar_pdf.get():
                pos = ReceiptPosition.TOP if self.recibo_pos.get() == "topo" else ReceiptPosition.BOTTOM
                danfe_cfg = DanfeConfig(receipt_pos=pos)
                dacte_cfg = DacteConfig()

            arquivos = sorted(entrada.rglob("*.xml"))
            total = len(arquivos)

            if total == 0:
                self._log_thread_safe("Nenhum arquivo XML encontrado.")
                self._progresso_thread_safe(0.0, "Nenhum XML encontrado.")
                self._fila_ui.put({"acao": "fim", "sucesso": True, "msg_final": "Nenhum XML encontrado na pasta origem."})
                return

            self._log_thread_safe(f"Encontrados {total} arquivo(s) XML.")
            
            ok = erro = pulado = 0
            regs_nfe: list[dict[str, Any]] = []
            regs_cte: list[dict[str, Any]] = []

            for i, xml_path in enumerate(arquivos, 1):
                relativo = xml_path.relative_to(entrada)
                tipo = detectar_tipo(xml_path)
                nome = xml_path.name

                frac = i / total
                self._progresso_thread_safe(frac, f"{i}/{total} — {nome}")

                if tipo == "desconhecido":
                    self._log_thread_safe(f"Ignorado: {nome} (tipo desconhecido)", "warning")
                    pulado += 1
                    continue

                if self.gerar_excel.get() and EXCEL_OK:
                    try:
                        if tipo == "nfe":
                            regs_nfe.append(extrair_nfe(xml_path))
                        else:
                            regs_cte.append(extrair_cte(xml_path))
                    except Exception as e:
                        self._log_thread_safe(f"Falha extração Excel: {nome} — {e}", "warning")

                if self.gerar_pdf.get() and FISCAL_OK:
                    pdf_path = (saida / relativo).with_suffix(".pdf")
                    pdf_path.parent.mkdir(parents=True, exist_ok=True)

                    if pdf_path.exists() and not self.sobrescrever.get():
                        self._log_thread_safe(f"Pulado PDF existente: {nome}")
                        pulado += 1
                        continue

                    try:
                        converter_pdf(xml_path, pdf_path, tipo, danfe_cfg, dacte_cfg)
                        self._log_thread_safe(f"Sucesso: {nome}")
                        ok += 1
                    except Exception as e:
                        self._log_thread_safe(f"Erro PDF {nome}: {e}", "error")
                        erro += 1
                else:
                    if tipo in ("nfe", "cte"):
                        ok += 1

            if self.gerar_excel.get() and EXCEL_OK and (regs_nfe or regs_cte):
                ts_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                nome_xl = f"notas_fiscais_{ts_str}.xlsx"
                caminho_xl = saida / nome_xl
                try:
                    self._log_thread_safe(f"Gerando Excel: {nome_xl}...")
                    gerar_excel(regs_nfe, regs_cte, caminho_xl)
                    self._log_thread_safe(f"Excel finalizado: {nome_xl}")
                except Exception as e:
                    self._log_thread_safe(f"Erro ao gerar Excel: {e}", "error")

            self._progresso_thread_safe(1.0, "Concluído!")
            msg = f"Processamento concluído.\nConvertidos: {ok}\nPulados: {pulado}\nErros: {erro}\n\nSaída: {saida.resolve()}"
            self._log_thread_safe("Processamento finalizado com sucesso.")
            self._fila_ui.put({"acao": "fim", "sucesso": True, "msg_final": msg})

        except Exception as e:
            logger.exception("Falha não tratada no worker de conversão.")
            self._log_thread_safe(f"ERRO FATAL: {e}", "error")
            self._fila_ui.put({"acao": "fim", "sucesso": False, "msg_final": f"Erro crítico na conversão:\n{e}"})

if __name__ == "__main__":
    app = AppConversorXML()
    app.mainloop()
