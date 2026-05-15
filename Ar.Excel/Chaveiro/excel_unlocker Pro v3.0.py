#!/usr/bin/env python3
"""
╔══════════════════════════════════════════════════════╗
║         Excel Unlocker Pro  v3.0                     ║
║  Gerenciamento profissional de proteções Excel        ║
╚══════════════════════════════════════════════════════╝

Suporte a:
  • Proteção de planilha (sheetProtection)  – hash legado XOR 16-bit e SHA-512
  • Proteção de estrutura (workbookProtection)
  • Criptografia de arquivo (AES-256 / RC4)  via msoffcrypto-tool
  • Quebra de senha de abertura por análise de criptografia + ataque inteligente

Dependências:
    pip install customtkinter msoffcrypto-tool openpyxl olefile
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
import time
import json
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

try:
    import olefile              # pip install olefile
    HAS_OLEFILE = True
except ImportError:
    HAS_OLEFILE = False

# ──────────────────────────────────────────────────────────────────────────────
#  Constantes / configuração visual
# ──────────────────────────────────────────────────────────────────────────────
APP_TITLE   = "Excel Unlocker Pro"
APP_VERSION = "3.0"
APP_W, APP_H = 820, 780

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
#  Núcleo da Opção 4 — Análise de criptografia de arquivo + ataque inteligente
# ══════════════════════════════════════════════════════════════════════════════

# ── Constantes de tipo de criptografia ───────────────────────────────────────
ENC_UNKNOWN  = "Desconhecido"
ENC_XOR      = "XOR/BIFF (legado — muito fraco, <1 ms)"
ENC_RC4_40   = "RC4 40-bit (legado — fraco)"
ENC_RC4_128  = "RC4 128-bit (legado — moderado)"
ENC_AES_128  = "AES-128 ECB Standard Encryption"
ENC_AES_192  = "AES-192 ECB Standard Encryption"
ENC_AES_256  = "AES-256 CBC Agile Encryption (PBKDF2-SHA512)"

# Mapeamento AlgID → nome legível
_ALG_MAP = {
    0x6801: "RC4",
    0x660E: "AES-128",
    0x660F: "AES-192",
    0x6610: "AES-256",
}
_HASH_MAP = {
    0x8004: "SHA-1",
    0x8003: "MD5",
    0x8009: "SHA-256",
    0x800C: "SHA-512",
}


def analyze_encryption(filepath: str | Path) -> dict:
    """
    Analisa o arquivo CFB (OLE) e extrai informações detalhadas sobre o
    esquema de criptografia usado pelo Excel.

    Retorna dict com chaves:
        enc_type, major, minor, alg, hash_alg, key_bits,
        salt_size, spin_count, provider, raw_xml (agile only)
    """
    info = {
        "enc_type":   ENC_UNKNOWN,
        "major":      0,
        "minor":      0,
        "alg":        "N/A",
        "hash_alg":   "N/A",
        "key_bits":   0,
        "salt_size":  0,
        "spin_count": 0,
        "provider":   "",
        "raw_xml":    "",
        "error":      "",
    }

    try:
        # ── Lê stream EncryptionInfo via olefile ─────────────────────────
        if HAS_OLEFILE:
            with olefile.OleFileIO(str(filepath)) as ole:
                if not ole.exists("EncryptionInfo"):
                    info["error"] = "Stream EncryptionInfo não encontrado"
                    return info
                raw = ole.openstream("EncryptionInfo").read()
        else:
            # Lê manualmente: CFB começa com assinatura 0xD0CF11E0...
            # Localiza EncryptionInfo via parsing básico (fallback)
            raw = _read_encryption_info_raw(filepath)
            if raw is None:
                info["error"] = "olefile não instalado e parsing manual falhou"
                return info

        if len(raw) < 8:
            info["error"] = "Stream EncryptionInfo muito curto"
            return info

        major, minor = struct.unpack_from("<HH", raw, 0)
        info["major"] = major
        info["minor"] = minor

        # ── Agile Encryption (major=4, minor=4) ─────────────────────────
        # 4 bytes reservados + XML
        if major == 4 and minor == 4:
            info["enc_type"] = ENC_AES_256
            xml_bytes = raw[8:]
            info["raw_xml"] = xml_bytes.decode("utf-8", errors="replace")
            _parse_agile_xml(xml_bytes, info)
            return info

        # ── Standard Encryption (major=3|4, minor=2|3) ───────────────────
        if minor in (2, 3):
            # 4 bytes flags + EncryptionHeader
            flags = struct.unpack_from("<I", raw, 4)[0]
            offset = 8
            hdr_size = struct.unpack_from("<I", raw, offset)[0]
            offset += 4  # skip HeaderSize field

            h_flags   = struct.unpack_from("<I", raw, offset)[0];     offset += 4
            _          = struct.unpack_from("<I", raw, offset)[0];     offset += 4  # SizeExtra
            alg_id    = struct.unpack_from("<I", raw, offset)[0];      offset += 4
            hash_id   = struct.unpack_from("<I", raw, offset)[0];      offset += 4
            key_bits  = struct.unpack_from("<I", raw, offset)[0];      offset += 4
            prov_type = struct.unpack_from("<I", raw, offset)[0];      offset += 4
            offset    += 8  # Reserved1 + Reserved2

            # CSPName (UTF-16LE, null-terminated)
            csp_raw = raw[offset: offset + hdr_size - 32]
            try:
                provider = csp_raw.rstrip(b"\x00").decode("utf-16-le")
            except Exception:
                provider = ""

            alg_name  = _ALG_MAP.get(alg_id,  f"AlgID=0x{alg_id:04X}")
            hash_name = _HASH_MAP.get(hash_id, f"HashID=0x{hash_id:04X}")

            info["alg"]       = alg_name
            info["hash_alg"]  = hash_name
            info["key_bits"]  = key_bits
            info["provider"]  = provider

            if alg_id == 0x6801:       # RC4
                info["enc_type"] = ENC_RC4_40 if key_bits <= 40 else ENC_RC4_128
            elif alg_id == 0x660E:
                info["enc_type"] = ENC_AES_128
            elif alg_id == 0x660F:
                info["enc_type"] = ENC_AES_192
            elif alg_id == 0x6610:
                info["enc_type"] = ENC_AES_256
            else:
                info["enc_type"] = f"Desconhecido (AlgID=0x{alg_id:04X})"

            # EncryptionVerifier — extrai salt
            ver_offset = 8 + 4 + hdr_size  # flags(4) + headerSize field(4) + header
            salt_size  = struct.unpack_from("<I", raw, ver_offset)[0]
            info["salt_size"] = salt_size
            return info

        # ── XOR Obfuscation (BIFF2-BIFF8, .xls antigo) ───────────────────
        if major == 1 and minor == 1:
            info["enc_type"] = ENC_XOR
            info["alg"]      = "XOR"
            info["key_bits"] = 16
            return info

    except Exception as e:
        info["error"] = str(e)

    return info


def _parse_agile_xml(xml_bytes: bytes, info: dict):
    """Extrai parâmetros do XML Agile Encryption."""
    try:
        root = ET.fromstring(xml_bytes)
        # Namespace pode variar
        for elem in root.iter():
            tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
            if tag == "encryptedKey":
                info["spin_count"] = int(elem.attrib.get("spinCount", 100000))
                info["hash_alg"]   = elem.attrib.get("hashAlgorithm", "SHA-512")
                info["key_bits"]   = int(elem.attrib.get("keyBits", 256))
                info["salt_size"]  = int(elem.attrib.get("saltSize", 16))
                info["alg"]        = elem.attrib.get("cipherAlgorithm", "AES")
                break
    except Exception:
        pass


def _read_encryption_info_raw(filepath) -> bytes | None:
    """
    Fallback mínimo para extrair stream EncryptionInfo de um CFB sem olefile.
    Lê apenas os primeiros 2 KB do arquivo para detectar versão.
    """
    try:
        with open(filepath, "rb") as f:
            # Assinatura CFB: D0 CF 11 E0 A1 B1 1A E1
            sig = f.read(8)
            if sig != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
                return None
            # Sem olefile não conseguimos navegar o FAT corretamente
            return None
    except Exception:
        return None


# ── Gerador inteligente de candidatos ────────────────────────────────────────

# Wordlist embutida — português + inglês + padrões universais
_BUILTIN_WORDLIST = [
    # vazios / triviais
    "", " ",
    # senhas numéricas universais
    "0", "1", "123", "1234", "12345", "123456", "1234567", "12345678",
    "0000", "0001", "1111", "2222", "3333", "4444", "5555", "6666",
    "7777", "8888", "9999", "00000", "11111", "99999", "000000", "111111",
    "999999", "123123", "321321", "112233", "998877",
    # senhas em português
    "senha", "senhas", "senha1", "senha123", "Senha1", "Senha123",
    "segredo", "acesso", "planilha", "excel", "arquivo", "dados",
    "empresa", "usuario", "admin", "administrador", "financeiro",
    "contabil", "contato", "relatorio", "projeto", "sistema",
    "protegido", "proteger", "bloqueado", "confidencial",
    "privado", "secreto", "chave", "codigo", "numero",
    # senhas em inglês
    "password", "pass", "passwd", "secret", "admin", "root",
    "login", "user", "access", "office", "excel123", "test",
    "welcome", "master", "manager", "default", "change",
    "letmein", "qwerty", "monkey", "dragon", "trustno1",
    "iloveyou", "sunshine", "princess", "football", "shadow",
    # padrões alfanuméricos comuns
    "abc123", "abc1234", "Abc123", "Abc1234", "ABC123",
    "pass123", "Pass123", "Pass@123", "P@ss123", "P@ssw0rd",
    "Admin1", "Admin123", "Admin@123", "root123", "Root123",
    "Welcome1", "Welcome123", "welcome1", "changeme", "Change1",
    # empresas / sistemas (comum em arquivos corporativos)
    "fiscal", "rh", "ti", "financas", "faturamento", "compras",
    "vendas", "logistica", "producao", "qualidade", "auditoria",
    # padrões com ano (gerado dinamicamente abaixo)
]

# Acrescenta anos dinamicamente: 1990–2030 e variações
for _y in range(1990, 2031):
    _BUILTIN_WORDLIST += [
        str(_y),
        f"senha{_y}", f"Senha{_y}", f"pass{_y}", f"excel{_y}",
        f"admin{_y}", f"fiscal{_y}", f"dados{_y}",
    ]

# Padrões sazonais/mês
_MONTHS_PT = ["jan","fev","mar","abr","mai","jun",
              "jul","ago","set","out","nov","dez"]
_MONTHS_EN = ["jan","feb","mar","apr","may","jun",
              "jul","aug","sep","oct","nov","dec"]
for _m in _MONTHS_PT + _MONTHS_EN:
    for _y in range(2018, 2031):
        _BUILTIN_WORDLIST.append(f"{_m}{_y}")
        _BUILTIN_WORDLIST.append(f"{_m.capitalize()}{_y}")

# Remove duplicatas preservando ordem
_seen = set()
_UNIQUE_WORDLIST = []
for _w in _BUILTIN_WORDLIST:
    if _w not in _seen:
        _seen.add(_w)
        _UNIQUE_WORDLIST.append(_w)
_BUILTIN_WORDLIST = _UNIQUE_WORDLIST


def generate_candidates(
    custom_wordlist: list[str] | None = None,
    wordlist_file: str | None = None,
    charset: str = "digits",
    max_brute_len: int = 4,
    stop_event: threading.Event | None = None,
):
    """
    Gerador que produz candidatos a senha em ordem de probabilidade:

      1. Wordlist embutida (português + inglês + padrões corporativos)
      2. Wordlist customizada passada como lista
      3. Wordlist de arquivo externo (.txt, um por linha)
      4. Força bruta crescente até max_brute_len caracteres
         charset: "digits" | "lower" | "alphanum" | "extended"
    """
    charsets = {
        "digits":   string.digits,
        "lower":    string.ascii_lowercase + string.digits,
        "alphanum": string.ascii_letters + string.digits,
        "extended": string.ascii_letters + string.digits + "!@#$%*-_.",
    }
    chars = charsets.get(charset, string.digits)

    # 1. Wordlist embutida
    for w in _BUILTIN_WORDLIST:
        if stop_event and stop_event.is_set():
            return
        yield w

    # 2. Wordlist customizada
    if custom_wordlist:
        for w in custom_wordlist:
            if stop_event and stop_event.is_set():
                return
            yield w.strip()

    # 3. Arquivo de wordlist
    if wordlist_file:
        try:
            with open(wordlist_file, "r", encoding="utf-8", errors="ignore") as fh:
                for line in fh:
                    if stop_event and stop_event.is_set():
                        return
                    w = line.strip()
                    if w:
                        yield w
        except Exception:
            pass

    # 4. Força bruta
    if max_brute_len > 0:
        for length in range(1, max_brute_len + 1):
            for combo in itertools.product(chars, repeat=length):
                if stop_event and stop_event.is_set():
                    return
                yield "".join(combo)




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

    # ══════════════════════════════════════════════════════════════════════════
    #  Opção 4 — Quebrar senha de abertura (criptografia de arquivo)
    # ══════════════════════════════════════════════════════════════════════════

    def crack_open_password(
        self,
        charset: str = "digits",
        max_brute_len: int = 4,
        wordlist_file: str | None = None,
        custom_words: list[str] | None = None,
        stop_event: threading.Event | None = None,
        status_cb=None,
    ) -> bool:
        """
        Opção 4: Tenta quebrar a senha de abertura de um arquivo Excel
        criptografado analisando o tipo de criptografia e testando
        candidatos gerados de forma inteligente.

        Fluxo:
          1. Verifica se o arquivo é realmente criptografado (CFB).
          2. Analisa o stream EncryptionInfo e loga todos os parâmetros.
          3. Para XOR/RC4-40 legados: tenta colisão de hash direto.
          4. Para AES: itera candidatos via msoffcrypto.OfficeFile.
          5. Ao encontrar: decifra, remove proteções internas e salva.
        """
        self._log("═" * 60)
        self._log(f"[OPÇÃO 4] Quebrar senha de abertura  →  {self.src.name}")
        self.progress(3)

        if not HAS_MSOFFCRYPTO:
            self._log(
                "msoffcrypto-tool não instalado. "
                "Execute: pip install msoffcrypto-tool", "ERROR"
            )
            return False

        if not self._is_encrypted():
            self._log(
                "O arquivo não está criptografado com senha de abertura. "
                "Para proteções de planilha use a Opção 3.", "WARNING"
            )
            return False

        # ── Etapa 1: Análise de criptografia ──────────────────────────────
        self._log("◈ Analisando estrutura de criptografia do arquivo...")
        enc = analyze_encryption(self.src)

        self._log(f"  Tipo       : {enc['enc_type']}")
        self._log(f"  Versão     : major={enc['major']} minor={enc['minor']}")
        self._log(f"  Algoritmo  : {enc['alg']}")
        self._log(f"  Hash       : {enc['hash_alg']}")
        self._log(f"  Chave      : {enc['key_bits']} bits")
        if enc["spin_count"]:
            self._log(f"  SpinCount  : {enc['spin_count']:,} iterações PBKDF2")
        if enc["salt_size"]:
            self._log(f"  Salt       : {enc['salt_size']} bytes")
        if enc["provider"]:
            self._log(f"  Provedor   : {enc['provider']}")
        if enc["error"]:
            self._log(f"  Aviso      : {enc['error']}", "WARNING")

        # Estimativa de velocidade
        if enc["spin_count"] >= 100000:
            self._log(
                "  ⚡ AES-256 + PBKDF2 com 100 k iterações: "
                "≈ 1–5 tentativas/s em Python puro. "
                "Wordlist embutida cobre senhas comuns e corporativas.", "WARNING"
            )
        elif "RC4" in enc["enc_type"] or "XOR" in enc["enc_type"]:
            self._log(
                "  ⚡ Criptografia legada RC4/XOR: muito rápida, "
                "brute force viável."
            )

        self.progress(8)

        # ── Etapa 2: Tentativa de quebra ──────────────────────────────────
        self._log("◈ Iniciando ataque inteligente de senhas...")
        self._log(
            f"  Charset    : {charset}  |  "
            f"Brute-force máx: {max_brute_len} chars  |  "
            f"Wordlist externa: {'Sim' if wordlist_file else 'Não'}"
        )

        found_password: str | None = None
        attempts      = 0
        t_start       = time.perf_counter()
        last_log_time = t_start

        # Fase 1: wordlist (80 % da barra de progresso)
        wordlist_total = len(_BUILTIN_WORDLIST) + (len(custom_words) if custom_words else 0)

        gen = generate_candidates(
            custom_wordlist=custom_words,
            wordlist_file=wordlist_file,
            charset=charset,
            max_brute_len=max_brute_len,
            stop_event=stop_event,
        )

        for candidate in gen:

            if stop_event and stop_event.is_set():
                self._log("◈ Operação interrompida pelo usuário.", "WARNING")
                return False

            attempts += 1

            # Tenta decifrar
            try:
                with open(self.src, "rb") as f:
                    office = msoffcrypto.OfficeFile(f)
                    office.load_key(password=candidate)
                    test_buf = io.BytesIO()
                    office.decrypt(test_buf)
                # Decifrou sem exceção → senha encontrada
                found_password = candidate
                break

            except Exception:
                pass  # Senha incorreta

            # Atualiza progresso e log periodicamente
            now = time.perf_counter()
            if now - last_log_time >= 3.0:
                elapsed  = now - t_start
                speed    = attempts / elapsed if elapsed > 0 else 0
                # Progresso aproximado pela posição na wordlist
                pct = min(8 + int(72 * min(attempts, wordlist_total)
                                  / max(wordlist_total, 1)), 80)
                self.progress(pct)
                self._log(
                    f"  Tentativas: {attempts:,}  |  "
                    f"Velocidade: {speed:.1f}/s  |  "
                    f"Última testada: '{candidate[:30]}'"
                )
                if status_cb:
                    status_cb(attempts, speed, candidate)
                last_log_time = now

        # ── Etapa 3: Resultado ────────────────────────────────────────────
        elapsed = time.perf_counter() - t_start
        self._log(
            f"◈ Busca encerrada — {attempts:,} tentativas em "
            f"{elapsed:.1f}s ({attempts/max(elapsed,0.001):.1f}/s)"
        )

        if found_password is None:
            self._log(
                "[RESULTADO] Senha não encontrada no espaço de busca atual.\n"
                "  Sugestões:\n"
                "  • Forneça um arquivo de wordlist personalizado\n"
                "  • Aumente o comprimento de força bruta\n"
                "  • Para AES-256 com senha complexa, use:\n"
                "    hashcat -m 9500 hash.txt wordlist.txt  (modo Office 2013+)\n"
                "    ou  john --format=office hash.txt",
                "ERROR",
            )
            self.progress(0)
            return False

        # Senha encontrada
        display = repr(found_password) if found_password else "'(vazia)'"
        self._log(f"[✓ SENHA ENCONTRADA] {display}", "SUCCESS" if False else "INFO")
        self._log(f"  Senha: {display}")
        self._log(f"  Tentativas até encontrar: {attempts:,}")
        self.progress(85)

        # ── Etapa 4: Decifra e salva ──────────────────────────────────────
        self._log("◈ Decifrando arquivo e removendo proteções internas...")
        try:
            with open(self.src, "rb") as f:
                office = msoffcrypto.OfficeFile(f)
                office.load_key(password=found_password)
                dec_buf = io.BytesIO()
                office.decrypt(dec_buf)

            self.progress(90)
            dec_buf.seek(0)

            # Remove também quaisquer proteções de planilha internas
            tmp_src  = self.src
            self.src = dec_buf  # type: ignore[assignment]
            out_path = self._output_path("SENHA_ENCONTRADA")
            ok       = self._remove_zip_protections(out_path)
            self.src = tmp_src

            if ok:
                self._log(
                    f"[SUCESSO] Arquivo desbloqueado salvo em: {out_path}"
                )
                # Salva a senha encontrada em arquivo de texto junto ao output
                pwd_file = out_path.parent / f"{out_path.stem}_SENHA.txt"
                pwd_file.write_text(
                    f"Arquivo: {self.src.name}\n"
                    f"Senha encontrada: {found_password}\n"
                    f"Tentativas: {attempts:,}\n"
                    f"Tempo: {elapsed:.1f}s\n"
                    f"Tipo de criptografia: {enc['enc_type']}\n",
                    encoding="utf-8",
                )
                self._log(f"  Senha registrada em: {pwd_file.name}")
            self.progress(100)
            return ok

        except Exception as e:
            self._log(f"Falha ao decifrar com senha encontrada: {e}", "ERROR")
            return False


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
        self._running    = False
        self._stop_event = threading.Event()
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
        if not HAS_OLEFILE:
            missing.append("olefile")

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
            self, width=APP_W - 32, height=220,
            fg_color=COLORS["panel"],
            segmented_button_fg_color=COLORS["accent"],
            segmented_button_selected_color=COLORS["green"],
            segmented_button_unselected_color=COLORS["accent"],
        )
        nb.pack(padx=16, pady=4)

        t1 = nb.add("  🗑  Opção 1 — Remover  ")
        t2 = nb.add("  🔄  Opção 2 — Trocar  ")
        t3 = nb.add("  🔬  Opção 3 — LookPic  ")
        t4 = nb.add("  🔑  Opção 4 — Quebrar abertura  ")

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
        self.t1_pwd = ctk.CTkEntry(t1, show="•", width=220,
                                   placeholder_text="deixe vazio se não houver")
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
            ("Senha atual:",           "t2_cur"),
            ("Nova senha:",            "t2_new"),
            ("Confirmar nova senha:",  "t2_cfm"),
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

        # ── Tab 4 ─────────────────────────────────────────────────────────
        # Descrição
        ctk.CTkLabel(
            t4,
            text=(
                "Identifica o tipo de criptografia, analisa os parâmetros e testa\n"
                "combinações inteligentes até encontrar a senha de abertura do arquivo."
            ),
            font=ctk.CTkFont(size=12),
            text_color=COLORS["text"],
            justify="left",
        ).grid(row=0, column=0, columnspan=4, sticky="w", padx=16, pady=(10, 6))

        # Charset
        ctk.CTkLabel(t4, text="Charset para força bruta:",
                     text_color=COLORS["subtext"],
                     font=ctk.CTkFont(size=11)).grid(
            row=1, column=0, sticky="e", padx=(16, 6), pady=3
        )
        self.t4_charset = ctk.CTkComboBox(
            t4,
            values=[
                "digits   (0-9)",
                "lower    (a-z + 0-9)",
                "alphanum (A-Za-z0-9)",
                "extended (A-Za-z0-9!@#$...)",
            ],
            width=220,
            state="readonly",
        )
        self.t4_charset.set("digits   (0-9)")
        self.t4_charset.grid(row=1, column=1, sticky="w", pady=3)

        # Max brute force length
        ctk.CTkLabel(t4, text="Comprimento máx (força bruta):",
                     text_color=COLORS["subtext"],
                     font=ctk.CTkFont(size=11)).grid(
            row=1, column=2, sticky="e", padx=(18, 6), pady=3
        )
        self.t4_maxlen = ctk.CTkComboBox(
            t4,
            values=["0 (desativado)", "1", "2", "3", "4", "5", "6"],
            width=130, state="readonly",
        )
        self.t4_maxlen.set("4")
        self.t4_maxlen.grid(row=1, column=3, sticky="w", pady=3)

        # Wordlist externa
        ctk.CTkLabel(t4, text="Wordlist externa (.txt):",
                     text_color=COLORS["subtext"],
                     font=ctk.CTkFont(size=11)).grid(
            row=2, column=0, sticky="e", padx=(16, 6), pady=3
        )
        self.t4_wl_var = ctk.StringVar(value="")
        ctk.CTkEntry(
            t4, textvariable=self.t4_wl_var,
            width=220, state="readonly",
            fg_color=COLORS["bg"], border_color=COLORS["accent"],
            placeholder_text="opcional",
        ).grid(row=2, column=1, sticky="w", pady=3)

        ctk.CTkButton(
            t4, text="Escolher", width=80,
            fg_color=COLORS["accent"],
            command=self._pick_wordlist,
        ).grid(row=2, column=2, sticky="w", padx=(4, 0), pady=3)

        # Velocidade estimada
        self.t4_speed_lbl = ctk.CTkLabel(
            t4,
            text="Aguardando execução...",
            font=ctk.CTkFont(family="Consolas", size=10),
            text_color=COLORS["subtext"],
        )
        self.t4_speed_lbl.grid(
            row=3, column=0, columnspan=4, sticky="w", padx=16, pady=(4, 0)
        )

    def _progress_section(self):
        frm = ctk.CTkFrame(self, fg_color=COLORS["panel"], corner_radius=10)
        frm.pack(fill="x", padx=16, pady=4)

        row = ctk.CTkFrame(frm, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=10)

        self.prog_bar = ctk.CTkProgressBar(
            row, width=480, height=18,
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
        self.prog_label.pack(side="left", padx=8)

        self.stop_btn = ctk.CTkButton(
            row, text="⏹  Parar", width=100, height=36,
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color=COLORS["red"],
            hover_color="#a52828",
            state="disabled",
            command=self._stop,
        )
        self.stop_btn.pack(side="right", padx=(6, 0))

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
            if self.dst_var.get() in ("", "Mesma pasta do arquivo"):
                self.dst_var.set(str(Path(path).parent))
            self._log_append(f"Arquivo selecionado: {path}")

    def _pick_dst(self):
        path = filedialog.askdirectory(title="Selecionar pasta de destino")
        if path:
            self.dst_var.set(path)

    def _pick_wordlist(self):
        path = filedialog.askopenfilename(
            title="Selecionar wordlist (.txt)",
            filetypes=[("Arquivo de texto", "*.txt"), ("Todos", "*.*")],
        )
        if path:
            self.t4_wl_var.set(path)
            self._log_append(f"Wordlist externa: {path}")

    def _stop(self):
        if self._running:
            self._stop_event.set()
            self._log_append("Sinal de parada enviado...", "WARNING")
            self.stop_btn.configure(state="disabled")

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
        elif "3" in tab:
            mode = "find"
        else:
            mode = "crack_open"

        self._running = True
        self._stop_event.clear()
        self.run_btn.configure(state="disabled", text="⏳  Processando...")
        self.stop_btn.configure(state="normal")
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
                0, lambda m=m, lvl=lvl: self._log_append(m, lvl)
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
                    ok = proc.change_password(current_pwd=cur, new_pwd=new)

            elif mode == "find":
                ok = proc.find_password()

            elif mode == "crack_open":
                # Lê parâmetros da Tab 4
                charset_raw = self.t4_charset.get().split()[0]   # ex: "digits"
                maxlen_raw  = self.t4_maxlen.get().split()[0]     # ex: "4" ou "0"
                try:
                    max_brute = int(maxlen_raw)
                except ValueError:
                    max_brute = 0

                wl_file = self.t4_wl_var.get() or None

                def status_cb(attempts, speed, last_pwd):
                    self.after(0, lambda: self.t4_speed_lbl.configure(
                        text=f"⚡ {attempts:,} tentativas  |  {speed:.1f}/s  "
                             f"|  última: '{last_pwd[:28]}'"
                    ))

                ok = proc.crack_open_password(
                    charset=charset_raw,
                    max_brute_len=max_brute,
                    wordlist_file=wl_file,
                    stop_event=self._stop_event,
                    status_cb=status_cb,
                )

        except Exception as e:
            logger.exception("Erro inesperado")
            self.after(0, lambda e=e: self._log_append(
                f"Erro inesperado: {e}", "ERROR"
            ))

        def _done():
            self._running = False
            self.run_btn.configure(state="normal", text="▶  Executar")
            self.stop_btn.configure(state="disabled")
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
    logger.info(f"msoffcrypto disponível : {HAS_MSOFFCRYPTO}")
    logger.info(f"openpyxl disponível   : {HAS_OPENPYXL}")
    logger.info(f"olefile disponível    : {HAS_OLEFILE}")
    logger.info("=" * 60)

    app = App()
    app.mainloop()

    logger.info(f"{APP_TITLE} — encerrado")


if __name__ == "__main__":
    main()
