# -*- coding: utf-8 -*-

"""
Lectura de metadatos de archivos WIMU (.QUI y .QUL).
"""

from __future__ import annotations

import sqlite3
import xml.etree.ElementTree as ET
from typing import Dict, Optional

from .utils import safe_text


def get_cfg_bytes(conn: sqlite3.Connection) -> Optional[bytes]:
    """Obtiene el archivo Sesion.cfg embebido en un .QUI."""
    try:
        row = conn.execute(
            """
            SELECT DATA
            FROM FILES
            WHERE lower(FILENAME)='sesion.cfg'
            LIMIT 1
            """
        ).fetchone()
    except sqlite3.Error:
        return None

    return row[0] if row else None


def parse_channel_metadata(cfg_bytes: Optional[bytes]) -> Dict[int, Dict[str, str]]:
    """
    Extrae los metadatos de los canales definidos en Sesion.cfg.
    """

    metadata: Dict[int, Dict[str, str]] = {}

    if not cfg_bytes:
        return metadata

    try:
        root = ET.fromstring(cfg_bytes.decode("utf-8", errors="replace"))
    except ET.ParseError:
        return metadata

    for channel in root.iter("CHANNEL"):

        row = {
            child.tag: safe_text(child.text)
            for child in channel
        }

        tag = row.get("TAG")

        if not tag:
            continue

        try:
            tag_number = int(float(tag))
        except ValueError:
            continue

        metadata[tag_number] = {
            "tag": str(tag_number),
            "code": row.get("CODE", ""),
            "name": row.get("NAME", ""),
            "magnitude": row.get("MAGNITUDE", ""),
            "unit": row.get("UNIT", ""),
            "frequency": row.get("FREC", ""),
            "series": row.get("SERIES", ""),
            "visible": row.get("VISIBLE", ""),
            "enabled": row.get("ENABLED", ""),
            "offset": row.get("OFFSET", ""),
            "scale": row.get("SCALE", ""),
            "time_offset": row.get("TIME_OFFSET", ""),
        }

    return metadata