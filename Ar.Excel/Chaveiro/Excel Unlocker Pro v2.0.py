#!/usr/bin/env python3
"""
╔══════════════════════════════════════════════════════╗
║         Excel Unlocker Pro  v2.0                     ║
║  Gerenciamento profissional de proteções Excel        ║
╚══════════════════════════════════════════════════════╝

Suporte a:
  • Proteção de planilha (sheetProtection)  – hash legado XOR 16-bit e SHA-512
  • Proteção de estrutura (workbookProtection)
  • Criptografia de arquivo (AES-256 / RC4)  via msoffcrypto-tool

Dependências:
    pip install customtkinter msoffcrypto-tool openpyxl
"""

# ──────────────────────────────────────────────────────────────────────────────
#  Imports
# ──────────────────────────────────────────────────────────────────────────────
import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox

import threading
import zipfile
import shutil
import os
import re
import io
import sys
import struct
import base64
import hashlib
import itertools
import logging
import string
from pathlib import Path
from datetime import datetime
from copy import deepcopy
import xml.etree.ElementTree as ET

# ── Dependências opcionais ──
try:
    import msoffcrypto          # pip install msoffcrypto-tool
    HAS_MSOFFCRYPTO = True
except ImportError:
    HAS_MSOFFCRYPTO = False

try:
    import openpyxl             # pip install openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# ──────────────────────────────────────────────────────────────────────────────
#  Constantes / configuração visual
# ──────────────────────────────────────────────────────────────────────────────
APP_TITLE   = "Excel Unlocker Pro"
APP_VERSION = "2.0"
APP_W, APP_H = 820, 720

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

COLORS = {
    "bg":       "#1a1a2e",
    "panel":    "#16213e",
    "accent":   "#0f3460",
    "green":    "#00b894",
    "yellow":   "#fdcb6e",
    "red":      "#d63031",
    "blue":     "#74b9ff",
    "text":     "#dfe6e9",
    "subtext":  "#636e72",
}

# ──────────────────────────────────────────────────────────────────────────────
#  Configuração de log
# ──────────────────────────────────────────────────────────────────────────────
LOG_DIR  = Path(__file__).resolve().parent
LOG_FILE = LOG_DIR / f"excel_unlocker_{datetime.now().strftime('%Y%m%d')}.log"

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)-8s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
    ],
)
logger = logging.getLogger("ExcelUnlocker")


# ══════════════════════════════════════════════════════════════════════════════
#  Núcleo criptográfico
# ══════════════════════════════════════════════════════════════════════════════

# ─────────────── Hash legado XOR (ECMA-376 §18.3.1.82) ───────────────────────

def _excel_legacy_hash(password: str) -> str:
    """
    Calcula o hash legado de 16-bit usado pela proteção de planilhas Excel
    (algoritmo XOR conforme ECMA-376 Part 4, §3.3.2.3).
    Retorna string hexadecimal maiúscula de 4 dígitos.
    """
    if not password:
        return "0000"
    chars = [ord(c) & 0xFF for c in password]
    h = 0
    for c in reversed(chars):
        # rotação circular 15-bit para a esquerda
        h = (((h >> 14) & 0x01) | ((h << 1) & 0x7FFF)) ^ c
    h = (((h >> 14) & 0x01) | ((h << 1) & 0x7FFF))
    h ^= len(chars)
    h ^= 0xCE4B
    return format(h & 0xFFFF, "04X")


def _find_collision_for_legacy_hash(target_hex: str,
                                    progress_cb=None) -> str | None:
    """
    Busca uma senha cujo hash XOR legado seja igual a target_hex.
    Estratégia: varre combinações de 1-4 caracteres imprimíveis.
    O espaço do hash é 16-bit (65 536 valores), garantindo colisão rápida.
    """
    target = target_hex.upper()
    charset = (
        string.ascii_letters
        + string.digits
        + string.punctuation
    )
    total_tested = 0

    for length in range(1, 5):
        for combo in itertools.product(charset, repeat=length):
            candidate = "".join(combo)
            total_tested += 1
            if _excel_legacy_hash(candidate) == target:
                logger.info(
                    "[HASH-COLISÃO] Senha encontrada: '%s' → hash %s "
                    "(testadas %d combinações)", candidate, target, total_tested
                )
                return candidate
            if progress_cb and total_tested % 5000 == 0:
                progress_cb(total_tested)
    return None


# ─────────────── Hash moderno SHA-512 (OOXML §3.3) ───────────────────────────

def _excel_modern_hash(password: str, salt_b64: str,
                       spin: int, algo: str = "SHA-512") -> str:
    """
    Calcula o hash OOXML moderno com iterações (SHA-1 / SHA-256 / SHA-512).
    Retorna base64 para comparação com o atributo hashValue do XML.
    """
    algo_map = {
        "SHA-512": hashlib.sha512,
        "SHA-256": hashlib.sha256,
        "SHA-1":   hashlib.sha1,
        "MD5":     hashlib.md5,
    }
    h_fn = algo_map.get(algo.upper(), hashlib.sha512)

    salt      = base64.b64decode(salt_b64)
    pwd_bytes = password.encode("utf-16-le")

    # H0 = hash(salt + password)
    h = h_fn(salt + pwd_bytes).digest()

    # Hn = hash(Hn-1 + iterador 4-byte little-endian)
    for i in range(spin):
        h = h_fn(h + struct.pack("<I", i)).digest()

    return base64.b64encode(h).decode()


def _crack_modern_hash(target_b64: str, salt_b64: str,
                       spin: int, algo: str,
                       progress_cb=None) -> str | None:
    """
    Tenta quebrar hash moderno via wordlist embutida + padrões simples.
    Para hashes SHA-512 com 100k iterações, cada tentativa leva ~50 ms;
    o volume de tentativas é limitado mas cobre senhas comuns.
    """
    # Wordlist embutida de senhas comuns
    WORDLIST = [
        "", "1234", "12345", "123456", "1234567", "12345678",
        "password", "senha", "excel", "admin", "secret", "pass",
        "planilha", "protegido", "proteger", "acesso", "login",
        "0000", "1111", "2222", "9999", "abc123", "qwerty",
        "teste", "test", "root", "user", "office", "microsoft",
        "abcd", "abcd1234", "senha123", "Pass@123", "P@ssw0rd",
        "Welcome1", "changeme", "letmein", "monkey", "dragon",
    ]
    # Gera padrões numéricos curtos
    for n in range(10000):
        WORDLIST.append(str(n))

    for i, pwd in enumerate(WORDLIST):
        h = _excel_modern_hash(pwd, salt_b64, spin, algo)
        if h == target_b64:
            logger.info("[SHA-CRACK] Senha encontrada: '%s' (tentativa %d)", pwd, i)
            return pwd
        if progress_cb and i % 100 == 0:
            progress_cb(i, len(WORDLIST))

    return None


# ══════════════════════════════════════════════════════════════════════════════
#  Processador de arquivos Excel
# ══════════════════════════════════════════════════════════════════════════════

NS_MAP = {
    "main":     "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r":        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc":       "http://schemas.openxmlformats.org/markup-compatibility/2006",
}

# Atributos de proteção que serão removidos / alterados
SHEET_PROT_ATTRS  = {
    "sheet", "objects", "scenarios", "formatCells", "formatColumns",
    "formatRows", "insertColumns", "insertRows", "insertHyperlinks",
    "deleteColumns", "deleteRows", "selectLockedCells", "sort",
    "autoFilter", "pivotTables", "selectUnlockedCells",
    "password", "hashValue", "saltValue", "spinCount", "algorithmName",
}

WB_PROT_ATTRS = {
    "lockStructure", "lockWindows", "lockRevision",
    "password", "hashValue", "saltValue", "spinCount", "algorithmName",
}


class ExcelProcessor:
    """
    Encapsula toda a lógica de manipulação do arquivo Excel.
    Opera diretamente no ZIP interno do .xlsx/.xlsm.
    """

    def __init__(self, src_path: str, dst_dir: str,
                 progress_cb=None, log_cb=None):
        self.src      = Path(src_path)
        self.dst_dir  = Path(dst_dir)
        self.progress = progress_cb or (lambda v, t=100: None)
        self.log      = log_cb      or (lambda m, lvl="INFO": None)
        self._stop    = threading.Event()

    # ──────────── helpers ────────────────────────────────────────────────────

    def _log(self, msg: str, level: str = "INFO"):
        getattr(logger, level.lower(), logger.info)(msg)
        self.log(msg, level)

    def _output_path(self, suffix: str) -> Path:
        stem = self.src.stem
        ext  = self.src.suffix
        return self.dst_dir / f"{stem}_{suffix}{ext}"

    def _is_encrypted(self) -> bool:
        """Verifica se o arquivo é um CFB criptografado (não é ZIP)."""
        with open(self.src, "rb") as f:
            sig = f.read(8)
        return sig == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"

    # ──────────── detecção de proteções no ZIP ───────────────────────────────

    def _scan_zip_protections(self) -> dict:
        """
        Lê o ZIP interno e retorna um dicionário com todas as proteções
        encontradas, junto com seus atributos XML.
        """
        result = {
            "encrypted_file": False,
            "workbook":       [],   # lista de atributos encontrados
            "sheets":         {},   # {sheet_part_name: [atributos]}
        }

        if self._is_encrypted():
            result["encrypted_file"] = True
            return result

        try:
            with zipfile.ZipFile(self.src, "r") as z:
                names = z.namelist()

                # workbook.xml
                wb_name = next(
                    (n for n in names if re.search(r"xl/workbook\.xml$", n)), None
                )
                if wb_name:
                    xml_bytes = z.read(wb_name)
                    root = ET.fromstring(xml_bytes)
                    for elem in root.iter():
                        tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
                        if tag == "workbookProtection":
                            result["workbook"] = dict(elem.attrib)

                # xl/worksheets/sheet*.xml
                sheet_names = [n for n in names
                               if re.match(r"xl/worksheets/sheet\d+\.xml$", n)]
                for sn in sheet_names:
                    xml_bytes = z.read(sn)
                    root = ET.fromstring(xml_bytes)
                    for elem in root.iter():
                        tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
                        if tag == "sheetProtection":
                            result["sheets"][sn] = dict(elem.attrib)
        except zipfile.BadZipFile as e:
            self._log(f"Erro ao abrir ZIP: {e}", "ERROR")

        return result

    # ──────────── manipulação XML ─────────────────────────────────────────────

    def _remove_protection_from_xml(self, xml_bytes: bytes,
                                    tag_name: str) -> bytes:
        """Remove todos os elementos com o tag_name especificado do XML."""
        # Preserva declaração XML e namespaces com string manipulation
        content = xml_bytes.decode("utf-8", errors="replace")

        # Remove o elemento inteiro (self-closing ou com filhos)
        pattern = re.compile(
            rf"<(?:[^:>]+:)?{re.escape(tag_name)}[^>]*/>"
            rf"|<(?:[^:>]+:)?{re.escape(tag_name)}[^>]*>.*?</(?:[^:>]+:)?{re.escape(tag_name)}>",
            re.DOTALL,
        )
        cleaned = pattern.sub("", content)
        return cleaned.encode("utf-8")

    def _update_protection_hash_in_xml(self, xml_bytes: bytes,
                                       tag_name: str,
                                       new_hash: str,
                                       attrs_to_keep: set,
                                       new_attrs: dict | None = None) -> bytes:
        """
        Substitui os atributos de hash do elemento de proteção pelo novo hash.
        Remove atributos relacionados a hash SHA e insere o legado ou novo.
        """
        content = xml_bytes.decode("utf-8", errors="replace")

        def replace_tag(m):
            tag_content = m.group(0)
            # Remove atributos de hash antigos
            for attr in ("password", "hashValue", "saltValue",
                         "spinCount", "algorithmName"):
                tag_content = re.sub(
                    rf'\s+{re.escape(attr)}="[^"]*"', "", tag_content
                )
            # Insere novo hash legado
            ins = f' password="{new_hash}"'
            if new_attrs:
                for k, v in new_attrs.items():
                    ins += f' {k}="{v}"'
            tag_content = re.sub(r"(\s*/?>)", ins + r"\1", tag_content, count=1)
            return tag_content

        pattern = re.compile(
            rf"<(?:[^:>]+:)?{re.escape(tag_name)}[^>]*/>"
            rf"|<(?:[^:>]+:)?{re.escape(tag_name)}[^>]*>",
        )
        result = pattern.sub(replace_tag, content)
        return result.encode("utf-8")

    # ──────────── reescrita de ZIP ────────────────────────────────────────────

    def _rewrite_zip(self, transforms: dict) -> bytes:
        """
        Recria o ZIP aplicando transforms = {part_name: new_bytes}.
        Retorna o conteúdo binário do novo ZIP.
        """
        buf = io.BytesIO()
        with zipfile.ZipFile(self.src, "r") as zin, \
             zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename in transforms:
                    data = transforms[item.filename]
                zout.writestr(item, data)
        return buf.getvalue()

    # ══════════════════════════════════════════════════════════════════════════
    #  Opção 1 — Remover senha
    # ══════════════════════════════════════════════════════════════════════════

    def remove_password(self, file_password: str = "") -> bool:
        """
        Remove proteções do arquivo Excel.
        • Arquivo criptografado: usa file_password para decifrar e salva sem criptografia.
        • Arquivo ZIP (não criptografado): remove elementos sheetProtection / workbookProtection.
        """
        self._log("═" * 60)
        self._log(f"[OPÇÃO 1] Remover senha  →  {self.src.name}")
        self.progress(5)

        out_path = self._output_path("SEM_SENHA")

        # ── Arquivo criptografado ──────────────────────────────────────────
        if self._is_encrypted():
            if not HAS_MSOFFCRYPTO:
                self._log(
                    "msoffcrypto-tool não instalado. "
                    "Execute: pip install msoffcrypto-tool", "ERROR"
                )
                return False
            if not file_password:
                self._log(
                    "Arquivo criptografado: informe a senha do arquivo.", "ERROR"
                )
                return False

            self._log("Arquivo criptografado detectado. Decifrando...")
            try:
                with open(self.src, "rb") as f:
                    office = msoffcrypto.OfficeFile(f)
                    office.load_key(password=file_password)
                    dec_buf = io.BytesIO()
                    office.decrypt(dec_buf)

                self.progress(50)
                dec_buf.seek(0)

                # Agora processa o ZIP resultante removendo proteções de planilha
                tmp_src  = self.src
                self.src = dec_buf  # type: ignore[assignment]
                ok = self._remove_zip_protections(out_path)
                self.src = tmp_src
                return ok
            except Exception as e:
                self._log(f"Falha ao decifrar: {e}", "ERROR")
                return False

        # ── Arquivo ZIP padrão ─────────────────────────────────────────────
        return self._remove_zip_protections(out_path)

    def _remove_zip_protections(self, out_path: Path) -> bool:
        prot = self._scan_zip_protections()
        transforms = {}

        try:
            with zipfile.ZipFile(
                self.src if isinstance(self.src, Path) else self.src,
                "r"
            ) as z:
                names = z.namelist()

                # workbook
                wb_name = next(
                    (n for n in names if re.search(r"xl/workbook\.xml$", n)), None
                )
                if wb_name and prot["workbook"]:
                    self._log(
                        f"Removendo workbookProtection de {wb_name}"
                        f"  attrs={list(prot['workbook'].keys())}"
                    )
                    data = z.read(wb_name)
                    transforms[wb_name] = self._remove_protection_from_xml(
                        data, "workbookProtection"
                    )

                # planilhas
                for sn, attrs in prot["sheets"].items():
                    self._log(
                        f"Removendo sheetProtection de {sn}"
                        f"  hash={attrs.get('password', attrs.get('hashValue', 'N/A'))}"
                    )
                    data = z.read(sn)
                    transforms[sn] = self._remove_protection_from_xml(
                        data, "sheetProtection"
                    )

                self.progress(70)

        except Exception as e:
            self._log(f"Erro ao processar ZIP: {e}", "ERROR")
            return False

        if not transforms:
            self._log(
                "Nenhuma proteção de planilha/pasta de trabalho encontrada no ZIP. "
                "Se o arquivo exige senha para abrir, use a opção com senha.",
                "WARNING",
            )

        new_zip = self._rewrite_zip(transforms)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_bytes(new_zip)
        self.progress(100)
        self._log(f"[SUCESSO] Arquivo salvo: {out_path}")
        return True

    # ══════════════════════════════════════════════════════════════════════════
    #  Opção 2 — Trocar senha
    # ══════════════════════════════════════════════════════════════════════════

    def change_password(self, current_pwd: str,
                        new_pwd: str,
                        change_file_pwd: bool = False) -> bool:
        """
        Troca a senha do arquivo.
        • change_file_pwd=True  → recifra o arquivo com nova senha (requer msoffcrypto)
        • change_file_pwd=False → atualiza hashes de proteção de planilha/pasta
        """
        self._log("═" * 60)
        self._log(f"[OPÇÃO 2] Trocar senha  →  {self.src.name}")
        self.progress(5)

        out_path = self._output_path("NOVA_SENHA")

        # ── Arquivo criptografado ──────────────────────────────────────────
        if self._is_encrypted():
            if not HAS_MSOFFCRYPTO:
                self._log("msoffcrypto-tool não instalado.", "ERROR")
                return False
            self._log("Decifrando com senha atual e recifrando com nova senha...")
            try:
                with open(self.src, "rb") as f:
                    office = msoffcrypto.OfficeFile(f)
                    office.load_key(password=current_pwd)
                    dec_buf = io.BytesIO()
                    office.decrypt(dec_buf)

                self.progress(40)
                dec_buf.seek(0)

                enc_buf = io.BytesIO()
                office2 = msoffcrypto.OfficeFile(dec_buf)
                office2.encrypt(new_pwd, enc_buf)

                self.progress(80)
                out_path.parent.mkdir(parents=True, exist_ok=True)
                out_path.write_bytes(enc_buf.getvalue())
                self.progress(100)
                self._log(f"[SUCESSO] Arquivo salvo: {out_path}")
                return True
            except Exception as e:
                self._log(f"Falha ao trocar senha do arquivo: {e}", "ERROR")
                return False

        # ── Proteção de planilha / pasta de trabalho ───────────────────────
        prot   = self._scan_zip_protections()
        new_h  = _excel_legacy_hash(new_pwd)
        transforms = {}

        try:
            with zipfile.ZipFile(self.src, "r") as z:
                names = z.namelist()

                wb_name = next(
                    (n for n in names if re.search(r"xl/workbook\.xml$", n)), None
                )
                if wb_name and prot["workbook"]:
                    self._log(
                        f"Trocando hash em {wb_name}  "
                        f"novo hash={new_h}"
                    )
                    data = z.read(wb_name)
                    transforms[wb_name] = self._update_protection_hash_in_xml(
                        data, "workbookProtection", new_h, WB_PROT_ATTRS
                    )

                for sn, attrs in prot["sheets"].items():
                    self._log(
                        f"Trocando hash em {sn}  "
                        f"novo hash={new_h}"
                    )
                    data = z.read(sn)
                    transforms[sn] = self._update_protection_hash_in_xml(
                        data, "sheetProtection", new_h, SHEET_PROT_ATTRS
                    )

                self.progress(70)
        except Exception as e:
            self._log(f"Erro ao processar ZIP: {e}", "ERROR")
            return False

        if not transforms:
            self._log(
                "Nenhuma proteção encontrada no ZIP para trocar.", "WARNING"
            )

        new_zip = self._rewrite_zip(transforms)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_bytes(new_zip)
        self.progress(100)
        self._log(f"[SUCESSO] Arquivo salvo: {out_path}")
        return True

    # ══════════════════════════════════════════════════════════════════════════
    #  Opção 3 — LookPic inteligente (encontrar senha por análise de hash)
    # ══════════════════════════════════════════════════════════════════════════

    def find_password(self) -> bool:
        """
        Analisa o algoritmo de hash usado pelo Excel e busca uma senha
        que gere o mesmo hash (colisão), desbloqueando a planilha.

        Para hashes legados (XOR 16-bit): colisão garantida, muito rápida.
        Para hashes modernos (SHA-512 / SHA-256): wordlist + padrões.
        """
        self._log("═" * 60)
        self._log(f"[OPÇÃO 3] LookPic — análise de hash  →  {self.src.name}")
        self.progress(5)

        if self._is_encrypted():
            self._log(
                "Arquivo criptografado: a senha de abertura não pode ser "
                "recuperada por análise de hash. "
                "Use força bruta externa ou forneça a senha atual.",
                "WARNING",
            )
            return False

        prot = self._scan_zip_protections()
        transforms = {}
        found_any  = False
        total_prot = len(prot["sheets"]) + (1 if prot["workbook"] else 0)

        if total_prot == 0:
            self._log("Nenhuma proteção encontrada no arquivo.", "WARNING")
            return False

        step = 80 / max(total_prot, 1)
        current_step = 10

        # ── helper: analisar um conjunto de atributos ─────────────────────
        def analyse_and_crack(attrs: dict, tag: str, part_name: str):
            nonlocal current_step, found_any

            algo       = attrs.get("algorithmName", "").upper()
            hash_val   = attrs.get("hashValue",  "")
            salt_val   = attrs.get("saltValue",  "")
            spin       = int(attrs.get("spinCount", 100000))
            legacy_pwd = attrs.get("password",   "")

            self._log(
                f"[ANÁLISE] {part_name}  "
                f"algoritmo={'LEGADO-XOR' if legacy_pwd else algo}  "
                f"hash={'(legado) ' + legacy_pwd if legacy_pwd else hash_val[:20] + '...'}"
            )

            # ── Caso 1: hash legado XOR 16-bit ────────────────────────────
            if legacy_pwd:
                self._log(
                    f"  → Hash XOR 16-bit detectado ({legacy_pwd}). "
                    "Buscando colisão no espaço de 65.536 valores..."
                )

                def prog(n):
                    pass  # progresso interno não bloqueia GUI

                found_pwd = _find_collision_for_legacy_hash(
                    legacy_pwd, progress_cb=prog
                )

                if found_pwd is not None:
                    self._log(
                        f"  ✓ Colisão encontrada! Senha equivalente: '{found_pwd}'"
                    )
                    # Substitui o hash pela senha encontrada (garantirá edição)
                    # Na prática, removemos a proteção diretamente
                    try:
                        with zipfile.ZipFile(self.src, "r") as z:
                            data = z.read(part_name)
                        transforms[part_name] = self._remove_protection_from_xml(
                            data, tag
                        )
                        found_any = True
                    except Exception as e:
                        self._log(f"  Erro ao aplicar remoção: {e}", "ERROR")
                else:
                    self._log(
                        "  ✗ Colisão não encontrada (improvável). "
                        "Removendo proteção diretamente...", "WARNING"
                    )
                    try:
                        with zipfile.ZipFile(self.src, "r") as z:
                            data = z.read(part_name)
                        transforms[part_name] = self._remove_protection_from_xml(
                            data, tag
                        )
                        found_any = True
                    except Exception as e:
                        self._log(f"  Erro: {e}", "ERROR")

            # ── Caso 2: hash moderno SHA-xxx ──────────────────────────────
            elif hash_val and salt_val:
                self._log(
                    f"  → Hash moderno {algo or 'SHA-512'} detectado. "
                    f"SpinCount={spin}. Testando wordlist..."
                )

                def prog2(i, total):
                    pct = int(current_step + step * (i / max(total, 1)))
                    self.progress(min(pct, 90))

                found_pwd = _crack_modern_hash(
                    hash_val, salt_val, spin,
                    algo or "SHA-512",
                    progress_cb=prog2,
                )

                if found_pwd is not None:
                    self._log(
                        f"  ✓ Senha encontrada via wordlist: '{found_pwd}'"
                    )
                    try:
                        with zipfile.ZipFile(self.src, "r") as z:
                            data = z.read(part_name)
                        transforms[part_name] = self._remove_protection_from_xml(
                            data, tag
                        )
                        found_any = True
                    except Exception as e:
                        self._log(f"  Erro ao aplicar remoção: {e}", "ERROR")
                else:
                    self._log(
                        "  ✗ Senha não encontrada na wordlist. "
                        "Hash SHA-512 com alta contagem de iterações requer "
                        "força bruta offline (ex.: hashcat). "
                        "Para uso imediato, forneça a senha atual e use a Opção 2.",
                        "WARNING",
                    )
            else:
                self._log(
                    "  Proteção presente mas sem atributos de hash reconhecidos. "
                    "Removendo elemento diretamente...", "WARNING"
                )
                try:
                    with zipfile.ZipFile(self.src, "r") as z:
                        data = z.read(part_name)
                    transforms[part_name] = self._remove_protection_from_xml(
                        data, tag
                    )
                    found_any = True
                except Exception as e:
                    self._log(f"  Erro: {e}", "ERROR")

            current_step = min(current_step + step, 90)
            self.progress(int(current_step))

        # ── Processa workbook ─────────────────────────────────────────────
        if prot["workbook"]:
            try:
                with zipfile.ZipFile(self.src, "r") as z:
                    wb_name = next(
                        n for n in z.namelist()
                        if re.search(r"xl/workbook\.xml$", n)
                    )
                analyse_and_crack(prot["workbook"], "workbookProtection", wb_name)
            except StopIteration:
                self._log("workbook.xml não encontrado.", "ERROR")

        # ── Processa planilhas ────────────────────────────────────────────
        for sn, attrs in prot["sheets"].items():
            analyse_and_crack(attrs, "sheetProtection", sn)

        if not found_any:
            self._log(
                "[RESULTADO] Não foi possível desbloquear automaticamente. "
                "Veja os detalhes acima.",
                "ERROR",
            )
            return False

        out_path = self._output_path("DESBLOQUEADO")
        new_zip  = self._rewrite_zip(transforms)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_bytes(new_zip)
        self.progress(100)
        self._log(f"[SUCESSO] Arquivo salvo: {out_path}")
        return True


# ══════════════════════════════════════════════════════════════════════════════
#  Interface Gráfica  (CustomTkinter)
# ══════════════════════════════════════════════════════════════════════════════

class App(ctk.CTk):

    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE}  v{APP_VERSION}")
        self.geometry(f"{APP_W}x{APP_H}")
        self.resizable(False, False)
        self.configure(fg_color=COLORS["bg"])
        self._running = False
        self._build_ui()

    # ──────────── construção da UI ────────────────────────────────────────────

    def _build_ui(self):
        self._header()
        self._dep_warning()
        self._file_section()
        self._option_tabs()
        self._progress_section()
        self._log_section()
        self._footer()

    def _header(self):
        frm = ctk.CTkFrame(self, fg_color=COLORS["accent"], corner_radius=0)
        frm.pack(fill="x")

        ctk.CTkLabel(
            frm,
            text=f"🔓  {APP_TITLE}",
            font=ctk.CTkFont(family="Segoe UI", size=22, weight="bold"),
            text_color=COLORS["blue"],
        ).pack(side="left", padx=20, pady=12)

        ctk.CTkLabel(
            frm,
            text=f"v{APP_VERSION}  •  XLSX / XLSM / XLS",
            font=ctk.CTkFont(size=11),
            text_color=COLORS["subtext"],
        ).pack(side="right", padx=20)

    def _dep_warning(self):
        missing = []
        if not HAS_MSOFFCRYPTO:
            missing.append("msoffcrypto-tool")
        if not HAS_OPENPYXL:
            missing.append("openpyxl")

        if missing:
            msg = "⚠  Dependências ausentes: " + ", ".join(missing)
            msg += "   →  pip install " + " ".join(missing)
            frm = ctk.CTkFrame(self, fg_color="#5a3a00", corner_radius=0)
            frm.pack(fill="x")
            ctk.CTkLabel(
                frm, text=msg,
                font=ctk.CTkFont(size=11),
                text_color=COLORS["yellow"],
            ).pack(padx=12, pady=5)

    def _file_section(self):
        frm = ctk.CTkFrame(self, fg_color=COLORS["panel"], corner_radius=10)
        frm.pack(fill="x", padx=16, pady=(12, 4))

        # ── Arquivo origem ────────────────────────────────────────────────
        row1 = ctk.CTkFrame(frm, fg_color="transparent")
        row1.pack(fill="x", padx=14, pady=(10, 4))

        ctk.CTkLabel(row1, text="📄  Arquivo Excel:",
                     font=ctk.CTkFont(size=12, weight="bold"),
                     text_color=COLORS["text"]).pack(side="left")

        self.src_var = ctk.StringVar(value="Nenhum arquivo selecionado")
        ctk.CTkEntry(
            row1, textvariable=self.src_var,
            width=480, state="readonly",
            fg_color=COLORS["bg"], border_color=COLORS["accent"],
        ).pack(side="left", padx=8)

        ctk.CTkButton(
            row1, text="Selecionar", width=100,
            fg_color=COLORS["accent"],
            command=self._pick_file,
        ).pack(side="left")

        # ── Pasta destino ─────────────────────────────────────────────────
        row2 = ctk.CTkFrame(frm, fg_color="transparent")
        row2.pack(fill="x", padx=14, pady=(0, 10))

        ctk.CTkLabel(row2, text="📁  Pasta destino: ",
                     font=ctk.CTkFont(size=12, weight="bold"),
                     text_color=COLORS["text"]).pack(side="left")

        self.dst_var = ctk.StringVar(value="Mesma pasta do arquivo")
        ctk.CTkEntry(
            row2, textvariable=self.dst_var,
            width=480, state="readonly",
            fg_color=COLORS["bg"], border_color=COLORS["accent"],
        ).pack(side="left", padx=8)

        ctk.CTkButton(
            row2, text="Selecionar", width=100,
            fg_color=COLORS["accent"],
            command=self._pick_dst,
        ).pack(side="left")

    def _option_tabs(self):
        self._opt_var = ctk.StringVar(value="1")

        nb = ctk.CTkTabview(
            self, width=APP_W - 32, height=200,
            fg_color=COLORS["panel"],
            segmented_button_fg_color=COLORS["accent"],
            segmented_button_selected_color=COLORS["green"],
            segmented_button_unselected_color=COLORS["accent"],
        )
        nb.pack(padx=16, pady=4)

        t1 = nb.add("  🗑  Opção 1 — Remover senha  ")
        t2 = nb.add("  🔄  Opção 2 — Trocar senha  ")
        t3 = nb.add("  🔬  Opção 3 — LookPic (encontrar)  ")

        self._nb = nb

        # ── Tab 1 ─────────────────────────────────────────────────────────
        ctk.CTkLabel(
            t1,
            text=(
                "Remove todas as proteções de planilha e pasta de trabalho.\n"
                "Para arquivos criptografados (protegidos com senha de abertura),\n"
                "informe a senha atual abaixo para decifrar o arquivo."
            ),
            font=ctk.CTkFont(size=12),
            text_color=COLORS["text"],
            justify="left",
        ).grid(row=0, column=0, columnspan=3, sticky="w", padx=16, pady=(12, 6))

        ctk.CTkLabel(t1, text="Senha do arquivo (se criptografado):",
                     text_color=COLORS["subtext"]).grid(
            row=1, column=0, sticky="e", padx=(16, 6), pady=4
        )
        self.t1_pwd = ctk.CTkEntry(t1, show="•", width=220, placeholder_text="deixe vazio se não houver")
        self.t1_pwd.grid(row=1, column=1, sticky="w", pady=4)

        # ── Tab 2 ─────────────────────────────────────────────────────────
        ctk.CTkLabel(
            t2,
            text=(
                "Localiza a senha atual e substitui pela nova.\n"
                "Funciona para proteção de planilha/pasta E para arquivos criptografados."
            ),
            font=ctk.CTkFont(size=12),
            text_color=COLORS["text"],
            justify="left",
        ).grid(row=0, column=0, columnspan=3, sticky="w", padx=16, pady=(12, 6))

        for i, (lbl, attr) in enumerate([
            ("Senha atual:",  "t2_cur"),
            ("Nova senha:",   "t2_new"),
            ("Confirmar nova senha:", "t2_cfm"),
        ], start=1):
            ctk.CTkLabel(t2, text=lbl, text_color=COLORS["subtext"]).grid(
                row=i, column=0, sticky="e", padx=(16, 6), pady=3
            )
            e = ctk.CTkEntry(t2, show="•", width=220)
            e.grid(row=i, column=1, sticky="w", pady=3)
            setattr(self, attr, e)

        # ── Tab 3 ─────────────────────────────────────────────────────────
        ctk.CTkLabel(
            t3,
            text=(
                "Analisa o algoritmo de hash utilizado pelo Excel para proteger planilhas.\n"
                "• Hash legado XOR 16-bit → colisão garantida em segundos.\n"
                "• Hash moderno SHA-512   → testa wordlist de senhas comuns.\n"
                "O arquivo desbloqueado é salvo na pasta destino."
            ),
            font=ctk.CTkFont(size=12),
            text_color=COLORS["text"],
            justify="left",
        ).pack(padx=16, pady=14, anchor="w")

    def _progress_section(self):
        frm = ctk.CTkFrame(self, fg_color=COLORS["panel"], corner_radius=10)
        frm.pack(fill="x", padx=16, pady=4)

        row = ctk.CTkFrame(frm, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=10)

        self.prog_bar = ctk.CTkProgressBar(
            row, width=540, height=18,
            progress_color=COLORS["green"],
            fg_color=COLORS["accent"],
        )
        self.prog_bar.pack(side="left")
        self.prog_bar.set(0)

        self.prog_label = ctk.CTkLabel(
            row, text="0 %",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=COLORS["green"], width=50,
        )
        self.prog_label.pack(side="left", padx=10)

        self.run_btn = ctk.CTkButton(
            row, text="▶  Executar", width=140, height=36,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color=COLORS["green"],
            hover_color="#00a381",
            command=self._run,
        )
        self.run_btn.pack(side="right")

    def _log_section(self):
        frm = ctk.CTkFrame(self, fg_color=COLORS["panel"], corner_radius=10)
        frm.pack(fill="both", expand=True, padx=16, pady=(4, 4))

        hdr = ctk.CTkFrame(frm, fg_color="transparent")
        hdr.pack(fill="x", padx=10, pady=(6, 0))

        ctk.CTkLabel(hdr, text="📋  Log de execução",
                     font=ctk.CTkFont(size=12, weight="bold"),
                     text_color=COLORS["text"]).pack(side="left")

        ctk.CTkButton(
            hdr, text="Limpar", width=70, height=26,
            fg_color=COLORS["accent"],
            command=self._clear_log,
        ).pack(side="right")

        self.log_box = ctk.CTkTextbox(
            frm, wrap="word",
            font=ctk.CTkFont(family="Consolas", size=11),
            fg_color=COLORS["bg"],
            text_color=COLORS["text"],
        )
        self.log_box.pack(fill="both", expand=True, padx=10, pady=(4, 10))
        self.log_box.configure(state="disabled")

    def _footer(self):
        frm = ctk.CTkFrame(self, fg_color=COLORS["accent"], corner_radius=0, height=28)
        frm.pack(fill="x", side="bottom")
        ctk.CTkLabel(
            frm,
            text=f"Log salvo em: {LOG_FILE}",
            font=ctk.CTkFont(size=10),
            text_color=COLORS["subtext"],
        ).pack(side="left", padx=12, pady=4)

    # ──────────── callbacks ───────────────────────────────────────────────────

    def _pick_file(self):
        path = filedialog.askopenfilename(
            title="Selecionar arquivo Excel",
            filetypes=[
                ("Excel files", "*.xlsx *.xlsm *.xls *.xlsb"),
                ("Todos os arquivos", "*.*"),
            ],
        )
        if path:
            self.src_var.set(path)
            # Define destino padrão como mesma pasta do arquivo
            if self.dst_var.get() in ("", "Mesma pasta do arquivo"):
                self.dst_var.set(str(Path(path).parent))
            self._log_append(f"Arquivo selecionado: {path}")

    def _pick_dst(self):
        path = filedialog.askdirectory(title="Selecionar pasta de destino")
        if path:
            self.dst_var.set(path)

    def _clear_log(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

    # ──────────── log na caixa de texto ──────────────────────────────────────

    def _log_append(self, msg: str, level: str = "INFO"):
        ts = datetime.now().strftime("%H:%M:%S")
        icons = {"INFO": "ℹ", "WARNING": "⚠", "ERROR": "✖", "SUCCESS": "✔"}
        colors_map = {
            "INFO":    COLORS["text"],
            "WARNING": COLORS["yellow"],
            "ERROR":   COLORS["red"],
            "SUCCESS": COLORS["green"],
        }
        icon = icons.get(level, "•")
        line = f"[{ts}] {icon}  {msg}\n"

        self.log_box.configure(state="normal")
        self.log_box.insert("end", line)
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    # ──────────── execução em thread ──────────────────────────────────────────

    def _run(self):
        if self._running:
            return

        src = self.src_var.get()
        dst = self.dst_var.get()

        if not src or src == "Nenhum arquivo selecionado":
            messagebox.showwarning("Atenção", "Selecione um arquivo Excel.")
            return
        if not Path(src).is_file():
            messagebox.showerror("Erro", f"Arquivo não encontrado:\n{src}")
            return
        if not dst or not Path(dst).exists():
            messagebox.showerror("Erro", "Pasta de destino inválida.")
            return

        tab = self._nb.get()
        if "1" in tab:
            mode = "remove"
        elif "2" in tab:
            mode = "change"
        else:
            mode = "find"

        self._running = True
        self.run_btn.configure(state="disabled", text="⏳  Processando...")
        self._set_progress(0)

        t = threading.Thread(
            target=self._worker,
            args=(src, dst, mode),
            daemon=True,
        )
        t.start()

    def _set_progress(self, val: int):
        """Atualiza barra e label de progresso (thread-safe via after)."""
        def _update():
            self.prog_bar.set(val / 100)
            self.prog_label.configure(text=f"{val} %")
        self.after(0, _update)

    def _worker(self, src: str, dst: str, mode: str):
        proc = ExcelProcessor(
            src, dst,
            progress_cb=lambda v, _=None: self._set_progress(v),
            log_cb=lambda m, lvl="INFO": self.after(
                0, lambda: self._log_append(m, lvl)
            ),
        )

        ok = False
        try:
            if mode == "remove":
                pwd = self.t1_pwd.get()
                ok = proc.remove_password(file_password=pwd)

            elif mode == "change":
                cur = self.t2_cur.get()
                new = self.t2_new.get()
                cfm = self.t2_cfm.get()
                if new != cfm:
                    self.after(0, lambda: self._log_append(
                        "Nova senha e confirmação não coincidem.", "ERROR"
                    ))
                else:
                    ok = proc.change_password(
                        current_pwd=cur, new_pwd=new
                    )

            elif mode == "find":
                ok = proc.find_password()

        except Exception as e:
            logger.exception("Erro inesperado")
            self.after(0, lambda: self._log_append(
                f"Erro inesperado: {e}", "ERROR"
            ))

        def _done():
            self._running = False
            self.run_btn.configure(state="normal", text="▶  Executar")
            if ok:
                self._log_append("Operação concluída com sucesso! ✔", "SUCCESS")
                self._set_progress(100)
            else:
                self._log_append("Operação concluída com falhas. Verifique o log.", "ERROR")

        self.after(0, _done)


# ══════════════════════════════════════════════════════════════════════════════
#  Entry-point
# ══════════════════════════════════════════════════════════════════════════════

def main():
    logger.info("=" * 60)
    logger.info(f"{APP_TITLE} v{APP_VERSION} — iniciado")
    logger.info(f"Python {sys.version}")
    logger.info(f"msoffcrypto disponível: {HAS_MSOFFCRYPTO}")
    logger.info(f"openpyxl disponível: {HAS_OPENPYXL}")
    logger.info("=" * 60)

    app = App()
    app.mainloop()

    logger.info(f"{APP_TITLE} — encerrado")


if __name__ == "__main__":
    main()