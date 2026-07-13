# -*- coding: utf-8 -*-

"""
Funciones auxiliares comunes.
"""

from __future__ import annotations

import math
import os
import re
import sys
import unicodedata
from pathlib import Path
from typing import BinaryIO, Iterable, Optional
from xml.sax.saxutils import escape


def pause_if_windows() -> None:
    """Evita que la consola se cierre al finalizar."""
    if os.name == "nt" and os.environ.get("WIMU_NO_PAUSE") != "1":
        input("\nProceso terminado. Pulsa Enter para cerrar...")


def clean_xml_text(value: object) -> str:
    text = "" if value is None else str(value)
    text = "".join(
        ch for ch in text
        if ch in "\t\n\r" or ord(ch) >= 32
    )
    return escape(text, {"\"": "&quot;"})


def safe_text(value: Optional[str]) -> str:
    return "" if value is None else str(value).strip()


def sanitize_excel_sheet_name(name: str, used: set[str]) -> str:
    name = unicodedata.normalize("NFKD", name)
    name = "".join(ch for ch in name if not unicodedata.combining(ch))
    name = re.sub(r"[\\/*?:\[\]]", "_", name)
    name = re.sub(r"\s+", "_", name.strip())
    name = name.strip("'") or "Hoja"

    base = name[:31]
    candidate = base
    counter = 2

    while candidate in used:
        suffix = f"_{counter}"
        candidate = base[:31-len(suffix)] + suffix
        counter += 1

    used.add(candidate)
    return candidate


def col_letter(col: int) -> str:
    result = ""

    while col:
        col, rem = divmod(col - 1, 26)
        result = chr(65 + rem) + result

    return result


def cell_ref(row: int, col: int) -> str:
    return f"{col_letter(col)}{row}"


def write_bytes(f: BinaryIO, text: str) -> None:
    f.write(text.encode("utf-8"))


def is_finite(value: float) -> bool:
    return math.isfinite(value)


def project_root() -> Path:
    return Path(__file__).resolve().parent.parent


def is_windows() -> bool:
    return sys.platform.startswith("win")