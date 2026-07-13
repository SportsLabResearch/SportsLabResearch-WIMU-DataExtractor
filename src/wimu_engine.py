
# -*- coding: utf-8 -*-
"""
Extrae datos de archivos WIMU/Quiko .QUL y .QUI a Excel en Windows.

Qué hace:
  - Busca una carpeta llamada "wimu" junto a este script.
  - Busca todos los archivos .qul y .qui dentro de esa carpeta.
  - Para cada archivo .qul/.qui, genera en esa misma carpeta un Excel llamado:
        <nombre_original>_resultados_extraidos.xlsx

Ejemplo:
  wimu\\Wimu_01.qui  ->  wimu\\Wimu_01_resultados_extraidos.xlsx

Uso en Windows:
  1) Instala Python desde https://www.python.org/downloads/ marcando "Add Python to PATH".
  2) Crea una carpeta llamada "wimu" junto a este script.
  3) Copia dentro de esa carpeta tu archivo .qul.
  4) Ejecuta este script con doble clic o desde CMD:
        py extraer_wimu_a_excel_windows.py

No necesita instalar pandas, openpyxl ni XlsxWriter: genera el .xlsx directamente.
"""

from __future__ import annotations

import math
import os
import re
import sys
import sqlite3
import struct
import collections
import platform
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from array import array
from datetime import datetime, timezone
from pathlib import Path
from typing import BinaryIO, Dict, Iterable, List, Optional, Tuple
from xml.sax.saxutils import escape

EXCEL_MAX_ROWS = 1_048_576
DATA_HEADER_ROW_EXCEL = 7        # fila visible de Excel
DATA_FIRST_ROW_EXCEL = 8         # primera fila de datos
ROWS_PER_SHEET = EXCEL_MAX_ROWS - DATA_FIRST_ROW_EXCEL + 1
INCLUIR_TIEMPO_SEGUNDOS = True   # cambia a False si quieres Excel más pequeño y rápido
ZIP_COMPRESSION_LEVEL = 1        # 1 = más rápido, 9 = más comprimido

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PACKAGE_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"


def pause_if_windows() -> None:
    """Evita que la ventana se cierre inmediatamente al hacer doble clic en Windows."""
    if os.name == "nt" and os.environ.get("WIMU_NO_PAUSE") != "1":
        input("\nProceso terminado. Pulsa Enter para cerrar...")


def find_wimu_folder() -> Path:
    """Devuelve la carpeta data/input del repositorio."""
    project_root = Path(__file__).resolve().parent.parent
    folder = project_root / "data" / "input"
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def output_folder() -> Path:
    """Devuelve la carpeta data/output del repositorio."""
    project_root = Path(__file__).resolve().parent.parent
    folder = project_root / "data" / "output"
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def clean_xml_text(value: object) -> str:
    """Elimina caracteres no válidos para XML y escapa texto."""
    text = "" if value is None else str(value)
    text = "".join(
        ch for ch in text
        if ch in "\t\n\r" or ord(ch) >= 32
    )
    return escape(text, {"\"": "&quot;"})


def safe_text(value: Optional[str]) -> str:
    return "" if value is None else str(value).strip()


def sanitize_excel_sheet_name(name: str, used: set) -> str:
    """Convierte texto en un nombre de hoja válido para Excel y único."""
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
        candidate = base[: 31 - len(suffix)] + suffix
        counter += 1
    used.add(candidate)
    return candidate


def col_letter(col_index_1_based: int) -> str:
    """Convierte número de columna 1=A, 2=B... a letra Excel."""
    result = ""
    n = col_index_1_based
    while n:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def cell_ref(row: int, col: int) -> str:
    return f"{col_letter(col)}{row}"


def inline_str_cell(row: int, col: int, value: object, style: int = 0) -> str:
    s_attr = f' s="{style}"' if style else ""
    return f'<c r="{cell_ref(row, col)}" t="inlineStr"{s_attr}><is><t>{clean_xml_text(value)}</t></is></c>'


def number_cell(row: int, col: int, value: object, style: int = 0) -> str:
    s_attr = f' s="{style}"' if style else ""
    try:
        if isinstance(value, float):
            if not math.isfinite(value):
                return f'<c r="{cell_ref(row, col)}"{s_attr}/>'
            val = format(value, ".12g")
        else:
            val = str(int(value))
    except Exception:
        return inline_str_cell(row, col, value, style)
    return f'<c r="{cell_ref(row, col)}"{s_attr}><v>{val}</v></c>'


def row_xml(row_num: int, cells: Iterable[str]) -> str:
    return f'<row r="{row_num}">' + "".join(cells) + '</row>\n'


def write_bytes(f: BinaryIO, text: str) -> None:
    f.write(text.encode("utf-8"))


def parse_channel_metadata(cfg_bytes: Optional[bytes]) -> Dict[int, Dict[str, str]]:
    """Lee Sesion.cfg para obtener nombres, unidades, frecuencia y metadatos de canales."""
    metadata: Dict[int, Dict[str, str]] = {}
    if not cfg_bytes:
        return metadata

    text = cfg_bytes.decode("utf-8", errors="replace")
    try:
        root = ET.fromstring(text)
    except ET.ParseError:
        return metadata

    for ch in root.iter("CHANNEL"):
        row = {child.tag: safe_text(child.text) for child in list(ch)}
        tag = row.get("TAG")
        if not tag:
            continue
        try:
            tag_num = int(float(tag))
        except ValueError:
            continue

        metadata[tag_num] = {
            "tag": str(tag_num),
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


def get_cfg_bytes(conn: sqlite3.Connection) -> Optional[bytes]:
    """Obtiene el archivo Sesion.cfg embebido dentro del .qui, si existe."""
    try:
        row = conn.execute(
            "SELECT DATA FROM FILES WHERE lower(FILENAME) = 'sesion.cfg' LIMIT 1"
        ).fetchone()
    except sqlite3.Error:
        return None
    return row[0] if row else None


def get_channel_tables(conn: sqlite3.Connection) -> List[Tuple[int, str]]:
    """Lista tablas CHANNELxx del archivo .qul o .qui."""
    rows = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name LIKE 'CHANNEL%'"
    ).fetchall()
    tables: List[Tuple[int, str]] = []
    for (name,) in rows:
        match = re.fullmatch(r"CHANNEL(\d+)", name)
        if match:
            tables.append((int(match.group(1)), name))
    return sorted(tables)


def bytes_to_numeric_array(blob: bytes, kind: str) -> array:
    """Convierte bytes little-endian en un array numérico nativo."""
    arr = array(kind)
    size = arr.itemsize
    usable = len(blob) - (len(blob) % size)
    arr.frombytes(blob[:usable])
    if sys.byteorder != "little":
        arr.byteswap()
    return arr


def decode_channel_blobs(x_blob: bytes, y_blob: bytes) -> Tuple[array, array, str]:
    """
    Decodifica los blobs X/Y de cada canal.

    En estos archivos WIMU .qui:
      - X suele venir como uint32 little-endian, normalmente tiempo en milisegundos.
      - Y suele venir como float32 little-endian, aunque se detecta float64 si corresponde.
    """
    x_values = bytes_to_numeric_array(x_blob, "I")
    n_x = len(x_values)

    if n_x == 0:
        return x_values, array("f"), "float32"

    bytes_per_y = len(y_blob) / n_x
    if abs(bytes_per_y - 8) < 0.01:
        y_values = bytes_to_numeric_array(y_blob, "d")
        y_type = "float64"
    elif abs(bytes_per_y - 4) < 0.01:
        y_values = bytes_to_numeric_array(y_blob, "f")
        y_type = "float32"
    else:
        y_values = bytes_to_numeric_array(y_blob, "f")
        y_type = "float32_detectado"

    n = min(len(x_values), len(y_values))
    if len(x_values) != n:
        del x_values[n:]
    if len(y_values) != n:
        del y_values[n:]
    return x_values, y_values, y_type


def channel_label(channel_num: int, meta: Dict[str, str]) -> str:
    """Genera un nombre entendible para la hoja del canal."""
    name = meta.get("name") or f"CANAL_{channel_num}"
    magnitude = meta.get("magnitude") or ""
    unit = meta.get("unit") or ""
    parts = [f"CH{channel_num}"]
    if magnitude:
        parts.append(magnitude)
    parts.append(name)
    if unit:
        parts.append(unit)
    return "_".join(parts)


def read_channel_info(conn: sqlite3.Connection, channel_num: int, table_name: str, meta: Dict[str, str]) -> Dict[str, object]:
    row = conn.execute(f"SELECT X, Y FROM {table_name} LIMIT 1").fetchone()
    if not row:
        x_blob, y_blob = b"", b""
    else:
        x_blob, y_blob = row[0] or b"", row[1] or b""
    x_values, y_values, y_type = decode_channel_blobs(x_blob, y_blob)
    sample_count = min(len(x_values), len(y_values))
    return {
        "channel_num": channel_num,
        "table_name": table_name,
        "meta": meta,
        "x_values": x_values,
        "y_values": y_values,
        "y_type": y_type,
        "sample_count": sample_count,
    }


def sheet_start(f: BinaryIO, freeze: bool = False, widths: Optional[List[Tuple[int, int, float]]] = None) -> None:
    write_bytes(f, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
    write_bytes(f, f'<worksheet xmlns="{NS_MAIN}" xmlns:r="{NS_REL}">\n')
    if freeze:
        write_bytes(f, '<sheetViews><sheetView workbookViewId="0"><pane ySplit="7" topLeftCell="A8" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>\n')
    if widths:
        write_bytes(f, '<cols>')
        for min_col, max_col, width in widths:
            write_bytes(f, f'<col min="{min_col}" max="{max_col}" width="{width}" customWidth="1"/>')
        write_bytes(f, '</cols>\n')
    write_bytes(f, '<sheetData>\n')


def sheet_end(f: BinaryIO) -> None:
    write_bytes(f, '</sheetData>\n</worksheet>')


def write_resumen_sheet(f: BinaryIO, qui_path: Path, channel_infos: List[Dict[str, object]], total_samples: int) -> None:
    sheet_start(f, freeze=False, widths=[(1, 1, 14), (2, 2, 28), (3, 9, 18)])

    write_bytes(f, row_xml(1, [inline_str_cell(1, 1, "Extracción WIMU .qui", 1)]))
    write_bytes(f, row_xml(3, [inline_str_cell(3, 1, "Archivo", 1), inline_str_cell(3, 2, qui_path.name)]))
    write_bytes(f, row_xml(4, [inline_str_cell(4, 1, "Ruta", 1), inline_str_cell(4, 2, str(qui_path))]))
    write_bytes(f, row_xml(5, [inline_str_cell(5, 1, "Canales", 1), number_cell(5, 2, len(channel_infos))]))
    write_bytes(f, row_xml(6, [inline_str_cell(6, 1, "Generado con", 1), inline_str_cell(6, 2, f"Python {platform.python_version()} / generador XLSX nativo")]))
    write_bytes(f, row_xml(7, [inline_str_cell(7, 1, "Total muestras", 1), number_cell(7, 2, total_samples)]))

    headers = ["Tabla", "TAG", "Código", "Nombre", "Magnitud", "Unidad", "Frecuencia", "Muestras", "Tipo Y"]
    write_bytes(f, row_xml(9, [inline_str_cell(9, i + 1, h, 1) for i, h in enumerate(headers)]))

    row_num = 10
    for info in channel_infos:
        meta = info["meta"]
        assert isinstance(meta, dict)
        channel_num = int(info["channel_num"])
        cells = [
            inline_str_cell(row_num, 1, info["table_name"]),
            inline_str_cell(row_num, 2, meta.get("tag", str(channel_num))),
            inline_str_cell(row_num, 3, meta.get("code", "")),
            inline_str_cell(row_num, 4, meta.get("name", "")),
            inline_str_cell(row_num, 5, meta.get("magnitude", "")),
            inline_str_cell(row_num, 6, meta.get("unit", "")),
            inline_str_cell(row_num, 7, meta.get("frequency", "")),
            number_cell(row_num, 8, int(info["sample_count"])),
            inline_str_cell(row_num, 9, info["y_type"]),
        ]
        write_bytes(f, row_xml(row_num, cells))
        row_num += 1

    sheet_end(f)


def write_channel_sheet_xml(f: BinaryIO, info: Dict[str, object]) -> None:
    channel_num = int(info["channel_num"])
    table_name = str(info["table_name"])
    meta = info["meta"]
    x_values = info["x_values"]
    y_values = info["y_values"]
    y_type = str(info["y_type"])
    sample_count = int(info["sample_count"])
    assert isinstance(meta, dict)
    assert isinstance(x_values, array)
    assert isinstance(y_values, array)

    sheet_start(f, freeze=True, widths=[(1, 1, 14), (2, 2, 14), (3, 3, 18), (4, 5, 16)])

    # Metadatos superiores.
    write_bytes(f, row_xml(1, [
        inline_str_cell(1, 1, "Tabla", 1), inline_str_cell(1, 2, table_name),
        inline_str_cell(1, 4, "Código", 1), inline_str_cell(1, 5, meta.get("code", "")),
    ]))
    write_bytes(f, row_xml(2, [
        inline_str_cell(2, 1, "Nombre", 1), inline_str_cell(2, 2, meta.get("name", "")),
        inline_str_cell(2, 4, "TAG", 1), inline_str_cell(2, 5, meta.get("tag", str(channel_num))),
    ]))
    write_bytes(f, row_xml(3, [
        inline_str_cell(3, 1, "Magnitud", 1), inline_str_cell(3, 2, meta.get("magnitude", "")),
        inline_str_cell(3, 4, "Tipo Y", 1), inline_str_cell(3, 5, y_type),
    ]))
    write_bytes(f, row_xml(4, [
        inline_str_cell(4, 1, "Unidad", 1), inline_str_cell(4, 2, meta.get("unit", "")),
        inline_str_cell(4, 4, "Muestras", 1), number_cell(4, 5, sample_count),
    ]))
    write_bytes(f, row_xml(5, [
        inline_str_cell(5, 1, "Frecuencia", 1), inline_str_cell(5, 2, meta.get("frequency", "")),
    ]))

    headers = ["tiempo_ms", "tiempo_s", "valor"] if INCLUIR_TIEMPO_SEGUNDOS else ["tiempo_ms", "valor"]
    write_bytes(f, row_xml(DATA_HEADER_ROW_EXCEL, [inline_str_cell(DATA_HEADER_ROW_EXCEL, i + 1, h, 1) for i, h in enumerate(headers)]))

    # Datos. Se escriben por bloques grandes para acelerar la generación del Excel.
    row_num = DATA_FIRST_ROW_EXCEL
    buffer: List[str] = []
    flush_every = 10_000

    if INCLUIR_TIEMPO_SEGUNDOS:
        for i in range(sample_count):
            t_ms = int(x_values[i])
            y = float(y_values[i])
            y_cell = f'<c r="C{row_num}"><v>{format(y, ".9g")}</v></c>' if math.isfinite(y) else f'<c r="C{row_num}"/>'
            buffer.append(
                f'<row r="{row_num}"><c r="A{row_num}"><v>{t_ms}</v></c>'
                f'<c r="B{row_num}"><v>{t_ms / 1000.0:.3f}</v></c>'
                f'{y_cell}</row>\n'
            )
            row_num += 1
            if len(buffer) >= flush_every:
                write_bytes(f, "".join(buffer))
                buffer.clear()
    else:
        for i in range(sample_count):
            t_ms = int(x_values[i])
            y = float(y_values[i])
            y_cell = f'<c r="B{row_num}"><v>{format(y, ".9g")}</v></c>' if math.isfinite(y) else f'<c r="B{row_num}"/>'
            buffer.append(
                f'<row r="{row_num}"><c r="A{row_num}"><v>{t_ms}</v></c>'
                f'{y_cell}</row>\n'
            )
            row_num += 1
            if len(buffer) >= flush_every:
                write_bytes(f, "".join(buffer))
                buffer.clear()

    if buffer:
        write_bytes(f, "".join(buffer))

    sheet_end(f)


def write_cfg_sheet_xml(f: BinaryIO, cfg_bytes: Optional[bytes]) -> None:
    sheet_start(f, freeze=False, widths=[(1, 1, 10), (2, 2, 120)])
    write_bytes(f, row_xml(1, [inline_str_cell(1, 1, "Línea", 1), inline_str_cell(1, 2, "Contenido", 1)]))
    if cfg_bytes:
        text = cfg_bytes.decode("utf-8", errors="replace")
        for idx, line in enumerate(text.splitlines(), start=2):
            write_bytes(f, row_xml(idx, [number_cell(idx, 1, idx - 1), inline_str_cell(idx, 2, line[:32767])]))
    sheet_end(f)


def content_types_xml(sheet_count: int) -> str:
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<Types xmlns="{NS_CONTENT_TYPES}">',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
    ]
    for i in range(1, sheet_count + 1):
        parts.append(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
    parts.append('</Types>')
    return "".join(parts)


def root_rels_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{NS_PACKAGE_REL}">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
        '</Relationships>'
    )


def workbook_xml(sheet_names: List[str]) -> str:
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<workbook xmlns="{NS_MAIN}" xmlns:r="{NS_REL}"><sheets>',
    ]
    for idx, name in enumerate(sheet_names, start=1):
        parts.append(f'<sheet name="{clean_xml_text(name)}" sheetId="{idx}" r:id="rId{idx}"/>')
    parts.append('</sheets></workbook>')
    return "".join(parts)


def workbook_rels_xml(sheet_count: int) -> str:
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<Relationships xmlns="{NS_PACKAGE_REL}">',
    ]
    for idx in range(1, sheet_count + 1):
        parts.append(f'<Relationship Id="rId{idx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{idx}.xml"/>')
    parts.append(f'<Relationship Id="rId{sheet_count + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>')
    parts.append('</Relationships>')
    return "".join(parts)


def styles_xml() -> str:
    # Estilos mínimos: 0 normal, 1 negrita para títulos/cabeceras.
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<styleSheet xmlns="{NS_MAIN}">'
        '<fonts count="2">'
        '<font><sz val="11"/><name val="Calibri"/></font>'
        '<font><b/><sz val="11"/><name val="Calibri"/></font>'
        '</fonts>'
        '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="2">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
        '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '</styleSheet>'
    )


def core_xml() -> str:
    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<dc:title>Resultados extraídos WIMU</dc:title>'
        '<dc:creator>Script Python</dc:creator>'
        f'<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>'
        f'<dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>'
        '</cp:coreProperties>'
    )


def app_xml(sheet_names: List[str]) -> str:
    sheets_vector = ''.join(f'<vt:lpstr>{clean_xml_text(name)}</vt:lpstr>' for name in sheet_names)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        '<Application>Python</Application>'
        '<DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>'
        f'<TitlesOfParts><vt:vector size="{len(sheet_names)}" baseType="lpstr">{sheets_vector}</vt:vector></TitlesOfParts>'
        f'<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>{len(sheet_names)}</vt:i4></vt:variant></vt:vector></HeadingPairs>'
        '</Properties>'
    )


def build_sheet_plan(channel_infos: List[Dict[str, object]], cfg_bytes: Optional[bytes]) -> Tuple[List[str], List[Tuple[str, Optional[Dict[str, object]]]]]:
    """Crea lista de nombres de hojas y qué debe escribirse en cada hoja."""
    used = set()
    sheet_names: List[str] = []
    plan: List[Tuple[str, Optional[Dict[str, object]]]] = []

    sheet_names.append(sanitize_excel_sheet_name("Resumen", used))
    plan.append(("resumen", None))

    for info in channel_infos:
        meta = info["meta"]
        assert isinstance(meta, dict)
        base_name = channel_label(int(info["channel_num"]), meta)
        total = int(info["sample_count"])
        parts = max(1, math.ceil(total / ROWS_PER_SHEET))
        if parts == 1:
            sheet_names.append(sanitize_excel_sheet_name(base_name, used))
            plan.append(("channel", info))
        else:
            # En caso de superar el límite de Excel, crea una hoja por tramo.
            # Para no duplicar datos en memoria, se guardan índices start/end en una copia ligera.
            for part in range(parts):
                start = part * ROWS_PER_SHEET
                end = min(start + ROWS_PER_SHEET, total)
                info_part = dict(info)
                info_part["part_start"] = start
                info_part["part_end"] = end
                sheet_names.append(sanitize_excel_sheet_name(f"{base_name}_p{part + 1}", used))
                plan.append(("channel_part", info_part))

    if cfg_bytes:
        sheet_names.append(sanitize_excel_sheet_name("Sesion_cfg", used))
        plan.append(("cfg", None))

    return sheet_names, plan


def write_channel_part_sheet_xml(f: BinaryIO, info: Dict[str, object]) -> None:
    """Escribe parte de un canal si supera el límite de filas. No se usa normalmente."""
    start = int(info.get("part_start", 0))
    end = int(info.get("part_end", int(info["sample_count"])))
    original_sample_count = int(info["sample_count"])
    info2 = dict(info)
    x_values = info["x_values"]
    y_values = info["y_values"]
    assert isinstance(x_values, array)
    assert isinstance(y_values, array)
    info2["x_values"] = x_values[start:end]
    info2["y_values"] = y_values[start:end]
    info2["sample_count"] = end - start
    write_channel_sheet_xml(f, info2)


def create_xlsx_native(output_path: Path, qui_path: Path, channel_infos: List[Dict[str, object]], cfg_bytes: Optional[bytes]) -> None:
    total_samples = sum(int(info["sample_count"]) for info in channel_infos)
    sheet_names, plan = build_sheet_plan(channel_infos, cfg_bytes)
    sheet_count = len(sheet_names)

    compression_kwargs = {"compression": zipfile.ZIP_DEFLATED, "compresslevel": ZIP_COMPRESSION_LEVEL}
    with zipfile.ZipFile(output_path, "w", allowZip64=True, **compression_kwargs) as z:
        z.writestr("[Content_Types].xml", content_types_xml(sheet_count))
        z.writestr("_rels/.rels", root_rels_xml())
        z.writestr("xl/workbook.xml", workbook_xml(sheet_names))
        z.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml(sheet_count))
        z.writestr("xl/styles.xml", styles_xml())
        z.writestr("docProps/core.xml", core_xml())
        z.writestr("docProps/app.xml", app_xml(sheet_names))

        for idx, (kind, info) in enumerate(plan, start=1):
            sheet_path = f"xl/worksheets/sheet{idx}.xml"
            with z.open(sheet_path, "w", force_zip64=True) as f:
                if kind == "resumen":
                    write_resumen_sheet(f, qui_path, channel_infos, total_samples)
                elif kind == "channel":
                    assert info is not None
                    write_channel_sheet_xml(f, info)
                elif kind == "channel_part":
                    assert info is not None
                    write_channel_part_sheet_xml(f, info)
                elif kind == "cfg":
                    write_cfg_sheet_xml(f, cfg_bytes)


def extract_one_qui(qui_path: Path) -> Path:
    """Extrae un archivo .qui y crea su Excel de resultados."""
    output_path = output_folder() / f"{qui_path.stem}_resultados_extraidos.xlsx"

    if output_path.exists():
        try:
            output_path.unlink()
        except PermissionError:
            raise PermissionError(
                f"No puedo sobrescribir '{output_path.name}'. Cierra ese Excel si está abierto y vuelve a ejecutar."
            )

    print(f"\nProcesando: {qui_path.name}")
    print(f"Salida:     {output_path.name}")

    conn = sqlite3.connect(str(qui_path))
    try:
        cfg_bytes = get_cfg_bytes(conn)
        channel_meta = parse_channel_metadata(cfg_bytes)
        channel_tables = get_channel_tables(conn)

        if not channel_tables:
            raise RuntimeError("No se encontraron tablas CHANNELxx dentro del archivo .qul o .qui.")

        channel_infos: List[Dict[str, object]] = []
        total_samples = 0
        for channel_num, table_name in channel_tables:
            meta = channel_meta.get(channel_num, {"tag": str(channel_num)})
            meta.setdefault("tag", str(channel_num))
            info = read_channel_info(conn, channel_num, table_name, meta)
            channel_infos.append(info)
            total_samples += int(info["sample_count"])
            nombre = meta.get("name", "") or f"Canal {channel_num}"
            magnitud = meta.get("magnitude", "")
            print(f"  - {table_name}: {nombre} / {magnitud} ({info['sample_count']} muestras)")

        print(f"  Total muestras: {total_samples}")
        print("  Creando Excel...")
        create_xlsx_native(output_path, qui_path, channel_infos, cfg_bytes)
    finally:
        conn.close()

    print(f"OK: {output_path}")
    return output_path



# ---------------------------------------------------------------------------
# Soporte adicional para archivos .QUL crudos de WIMU.
# ---------------------------------------------------------------------------

QUL_MAGIC = b"\xe1\xed"
QUL_EXTENSIONS = (".qul", ".qui")

# En los .qul observados, estos códigos de paquete son datos procesados que
# apuntan a sensores definidos en el XML con otro código.
QUL_ALIAS_SENSOR_CODE = {
    44: 300,  # ACEL procesado: X, Y, Z
    45: 301,  # ATTITUDE: 12 canales
    46: 302,  # GYRO procesado: X, Y, Z
}

# Orden amigable de hojas para los sensores habituales.
QUL_ORDER = {
    (1, 81): 10,  # GPS
    (6, 44): 20,  # ACEL
    (6, 46): 30,  # GYRO
    (6, 45): 40,  # ATTITUDE
    (6, 25): 50,  # MAG
    (6, 31): 60,  # BAR
    (6, 62): 70,  # HRM
    (6, 41): 80,  # STATUS
    (6, 42): 90,  # Battery
}


def parse_qul_sensors(xml_bytes: bytes) -> Dict[int, Dict[str, object]]:
    """Lee el XML inicial de un .qul y devuelve metadatos por código de sensor."""
    sensors: Dict[int, Dict[str, object]] = {}
    try:
        root = ET.fromstring(xml_bytes.decode("utf-8", errors="replace"))
    except ET.ParseError:
        return sensors

    for sensor in root.findall(".//SENSOR"):
        code_text = safe_text(sensor.findtext("CODE"))
        try:
            code = int(float(code_text))
        except ValueError:
            continue

        channels: List[Dict[str, str]] = []
        for ch in sensor.findall("./CHANNELS/CHANNEL"):
            channels.append({
                "name": safe_text(ch.findtext("NAME")) or "valor",
                "code": safe_text(ch.findtext("CODE")),
                "scale": safe_text(ch.findtext("SCALE")),
                "offset": safe_text(ch.findtext("OFFSET")),
            })

        sensors[code] = {
            "sensor_code": code,
            "name": safe_text(sensor.findtext("NAME")) or f"SENSOR_{code}",
            "type": safe_text(sensor.findtext("TYPE")),
            "enabled": safe_text(sensor.findtext("ENABLED")),
            "visible": safe_text(sensor.findtext("VISIBLE")),
            "magnitude": safe_text(sensor.findtext("MAGNITUDE")),
            "unit": safe_text(sensor.findtext("UNIT")),
            "channels": channels,
        }
    return sensors


def parse_qul_container(qul_path: Path) -> Dict[str, object]:
    """Carga un .qul, separa XML/metadatos y localiza el inicio binario."""
    raw = qul_path.read_bytes()
    marker = b"</NODE>"
    xml_end = raw.find(marker)
    if xml_end < 0:
        raise RuntimeError("No encontré el bloque XML inicial </NODE> del archivo .qul.")
    xml_end += len(marker)
    xml_bytes = raw[:xml_end]

    data_start = raw.find(QUL_MAGIC, xml_end)
    if data_start < 0:
        raise RuntimeError("No encontré el inicio de paquetes binarios del .qul.")

    tail_bytes = raw[xml_end:data_start]
    tail_text = tail_bytes.decode("utf-8", errors="replace")
    time_meta: Dict[str, str] = {}
    for tag in ("TIMEZ", "TIMEO", "TIMEU"):
        m = re.search(rf"<{tag}>(.*?)</{tag}>", tail_text, flags=re.DOTALL)
        if m:
            time_meta[tag] = m.group(1).strip()

    return {
        "raw": raw,
        "xml_bytes": xml_bytes,
        "tail_text": tail_text,
        "time_meta": time_meta,
        "data_start": data_start,
        "sensors": parse_qul_sensors(xml_bytes),
    }


def iter_qul_packets(raw: bytes, data_start: int):
    """
    Itera paquetes .qul.

    Estructura observada:
      magic 2 bytes  = E1 ED
      size  2 bytes  = tamaño del cuerpo
      type  1 byte   = tipo de paquete
      body  size bytes

    Para paquetes de datos, el cuerpo suele ser:
      raw_code 1 byte, flag 1 byte, tiempo_ms uint32, valores..., checksum 1 byte.
    """
    o = data_start
    n = len(raw)
    while o + 5 <= n:
        if raw[o:o + 2] != QUL_MAGIC:
            nearby = raw[o:o + 24].hex(" ")
            raise RuntimeError(f"Paquete .qul desalineado en byte {o}. Bytes: {nearby}")
        size = struct.unpack_from("<H", raw, o + 2)[0]
        packet_type = raw[o + 4]
        body_start = o + 5
        body_end = body_start + size
        if body_end > n:
            raise RuntimeError(f"Paquete .qul truncado en byte {o}; tamaño declarado {size}.")
        body = raw[body_start:body_end]
        yield o, packet_type, size, body
        o = body_end

    trailing = raw[o:]
    if trailing and any(b != 0 for b in trailing):
        raise RuntimeError(f"Quedaron {len(trailing)} bytes sin interpretar al final del .qul.")


def body_code_flag(body: bytes) -> Tuple[int, int]:
    code = body[0] if len(body) >= 1 else -1
    flag = body[1] if len(body) >= 2 else -1
    return code, flag


def collect_qul_counts(raw: bytes, data_start: int) -> Dict[Tuple[int, int, int, int], int]:
    counts: Dict[Tuple[int, int, int, int], int] = collections.Counter()
    for _, packet_type, size, body in iter_qul_packets(raw, data_start):
        code, flag = body_code_flag(body)
        counts[(packet_type, code, flag, size)] += 1
    return dict(counts)


def get_qul_sensor_for_packet(packet_type: int, raw_code: int, sensors: Dict[int, Dict[str, object]]) -> Dict[str, object]:
    sensor_code = QUL_ALIAS_SENSOR_CODE.get(raw_code, raw_code)
    sensor = sensors.get(sensor_code)
    if sensor:
        return sensor
    return {
        "sensor_code": sensor_code,
        "name": f"PACKET_{packet_type}_{raw_code}",
        "type": "",
        "enabled": "",
        "visible": "",
        "magnitude": "",
        "unit": "",
        "channels": [],
    }


def get_qul_decoder(packet_type: int, raw_code: int) -> str:
    if packet_type == 1 and raw_code == 81:
        return "float64"
    if packet_type == 6:
        return "float32"
    if packet_type == 2:
        return "time_sync"
    return "bytes"


def get_qul_channel_names(sensor: Dict[str, object], value_count: int) -> List[str]:
    raw_channels = sensor.get("channels", [])
    names: List[str] = []
    if isinstance(raw_channels, list):
        for ch in raw_channels:
            if isinstance(ch, dict):
                names.append(safe_text(ch.get("name")) or f"valor_{len(names) + 1}")
    while len(names) < value_count:
        names.append(f"valor_{len(names) + 1}")
    return names[:value_count]


def infer_qul_value_count(packet_type: int, raw_code: int, size: int) -> int:
    # body: raw_code + flag + timestamp_ms + payload + checksum
    payload_len = max(0, size - 7)
    decoder = get_qul_decoder(packet_type, raw_code)
    if decoder == "float64":
        return payload_len // 8
    if decoder == "float32":
        return payload_len // 4
    return payload_len


def decode_qul_values(body: bytes, decoder: str) -> Tuple[float, ...]:
    if len(body) < 7:
        return tuple()
    payload = body[6:-1]
    if decoder == "float64":
        n = len(payload) // 8
        return struct.unpack("<" + "d" * n, payload[: n * 8]) if n else tuple()
    if decoder == "float32":
        n = len(payload) // 4
        return struct.unpack("<" + "f" * n, payload[: n * 4]) if n else tuple()
    return tuple(float(x) for x in payload)


def timestamp_ms_from_body(body: bytes) -> int:
    if len(body) < 6:
        return 0
    return struct.unpack_from("<I", body, 2)[0]


def build_qul_packet_infos(counts: Dict[Tuple[int, int, int, int], int], sensors: Dict[int, Dict[str, object]]) -> List[Dict[str, object]]:
    infos: List[Dict[str, object]] = []
    for (packet_type, raw_code, flag, size), count in counts.items():
        if packet_type == 2:
            continue
        decoder = get_qul_decoder(packet_type, raw_code)
        if decoder == "bytes":
            # Mantiene el script seguro ante paquetes desconocidos sin romper la extracción principal.
            continue
        value_count = infer_qul_value_count(packet_type, raw_code, size)
        sensor = get_qul_sensor_for_packet(packet_type, raw_code, sensors)
        channel_names = get_qul_channel_names(sensor, value_count)
        infos.append({
            "packet_type": packet_type,
            "raw_code": raw_code,
            "flag": flag,
            "size": size,
            "count": count,
            "decoder": decoder,
            "value_count": value_count,
            "sensor": sensor,
            "channel_names": channel_names,
        })

    def sort_key(info: Dict[str, object]) -> Tuple[int, int, int, int]:
        packet_type = int(info["packet_type"])
        raw_code = int(info["raw_code"])
        return (QUL_ORDER.get((packet_type, raw_code), 999), packet_type, raw_code, int(info["size"]))

    return sorted(infos, key=sort_key)


def packet_sheet_base_name(info: Dict[str, object]) -> str:
    sensor = info["sensor"]
    assert isinstance(sensor, dict)
    name = safe_text(sensor.get("name")) or f"PACKET_{info['packet_type']}_{info['raw_code']}"
    return f"{name}_{info['raw_code']}"


def write_qul_resumen_sheet(
    f: BinaryIO,
    qul_path: Path,
    packet_infos: List[Dict[str, object]],
    counts: Dict[Tuple[int, int, int, int], int],
    time_meta: Dict[str, str],
    data_start: int,
) -> None:
    sheet_start(f, freeze=False, widths=[(1, 1, 16), (2, 2, 30), (3, 12, 18)])

    total_packets = sum(counts.values())
    total_samples = sum(int(info["count"]) for info in packet_infos)

    write_bytes(f, row_xml(1, [inline_str_cell(1, 1, "Extracción WIMU .qul", 1)]))
    write_bytes(f, row_xml(3, [inline_str_cell(3, 1, "Archivo", 1), inline_str_cell(3, 2, qul_path.name)]))
    write_bytes(f, row_xml(4, [inline_str_cell(4, 1, "Ruta", 1), inline_str_cell(4, 2, str(qul_path))]))
    write_bytes(f, row_xml(5, [inline_str_cell(5, 1, "Formato detectado", 1), inline_str_cell(5, 2, "WIMU .qul raw logger")]))
    write_bytes(f, row_xml(6, [inline_str_cell(6, 1, "Paquetes totales", 1), number_cell(6, 2, total_packets)]))
    write_bytes(f, row_xml(7, [inline_str_cell(7, 1, "Muestras exportadas", 1), number_cell(7, 2, total_samples)]))
    write_bytes(f, row_xml(8, [inline_str_cell(8, 1, "Inicio binario byte", 1), number_cell(8, 2, data_start)]))
    write_bytes(f, row_xml(9, [inline_str_cell(9, 1, "Zona horaria", 1), inline_str_cell(9, 2, time_meta.get("TIMEZ", ""))]))
    write_bytes(f, row_xml(10, [inline_str_cell(10, 1, "Offset horario", 1), inline_str_cell(10, 2, time_meta.get("TIMEO", ""))]))
    write_bytes(f, row_xml(11, [inline_str_cell(11, 1, "TIMEU", 1), inline_str_cell(11, 2, time_meta.get("TIMEU", ""))]))
    write_bytes(f, row_xml(12, [inline_str_cell(12, 1, "Generado con", 1), inline_str_cell(12, 2, f"Python {platform.python_version()} / generador XLSX nativo")]))

    headers = ["Tipo", "Código raw", "Flag", "Tamaño", "Sensor", "Código sensor", "Magnitud", "Unidad", "Canales", "Muestras", "Decodificador"]
    write_bytes(f, row_xml(14, [inline_str_cell(14, i + 1, h, 1) for i, h in enumerate(headers)]))

    row_num = 15
    for info in packet_infos:
        sensor = info["sensor"]
        assert isinstance(sensor, dict)
        channel_names = info["channel_names"]
        assert isinstance(channel_names, list)
        cells = [
            number_cell(row_num, 1, int(info["packet_type"])),
            number_cell(row_num, 2, int(info["raw_code"])),
            number_cell(row_num, 3, int(info["flag"])),
            number_cell(row_num, 4, int(info["size"])),
            inline_str_cell(row_num, 5, sensor.get("name", "")),
            number_cell(row_num, 6, int(sensor.get("sensor_code", info["raw_code"]))),
            inline_str_cell(row_num, 7, sensor.get("magnitude", "")),
            inline_str_cell(row_num, 8, sensor.get("unit", "")),
            inline_str_cell(row_num, 9, ", ".join(str(x) for x in channel_names)),
            number_cell(row_num, 10, int(info["count"])),
            inline_str_cell(row_num, 11, info["decoder"]),
        ]
        write_bytes(f, row_xml(row_num, cells))
        row_num += 1

    # Tabla de todos los paquetes, incluidos sincronización o desconocidos.
    row_num += 2
    write_bytes(f, row_xml(row_num, [inline_str_cell(row_num, 1, "Conteo completo de paquetes", 1)]))
    row_num += 1
    headers2 = ["Tipo", "Código raw", "Flag", "Tamaño", "Conteo"]
    write_bytes(f, row_xml(row_num, [inline_str_cell(row_num, i + 1, h, 1) for i, h in enumerate(headers2)]))
    row_num += 1
    for key in sorted(counts):
        packet_type, raw_code, flag, size = key
        cells = [
            number_cell(row_num, 1, packet_type),
            number_cell(row_num, 2, raw_code),
            number_cell(row_num, 3, flag),
            number_cell(row_num, 4, size),
            number_cell(row_num, 5, counts[key]),
        ]
        write_bytes(f, row_xml(row_num, cells))
        row_num += 1

    sheet_end(f)


def write_qul_packet_sheet_xml(f: BinaryIO, info: Dict[str, object], raw: bytes, data_start: int) -> None:
    packet_type = int(info["packet_type"])
    raw_code = int(info["raw_code"])
    flag = int(info["flag"])
    size = int(info["size"])
    count = int(info["count"])
    decoder = str(info["decoder"])
    sensor = info["sensor"]
    channel_names = info["channel_names"]
    assert isinstance(sensor, dict)
    assert isinstance(channel_names, list)

    end_col = 2 + len(channel_names)
    widths = [(1, 1, 14), (2, 2, 14), (3, max(3, end_col), 16)]
    sheet_start(f, freeze=True, widths=widths)

    write_bytes(f, row_xml(1, [
        inline_str_cell(1, 1, "Sensor", 1), inline_str_cell(1, 2, sensor.get("name", "")),
        inline_str_cell(1, 4, "Código raw", 1), number_cell(1, 5, raw_code),
    ]))
    write_bytes(f, row_xml(2, [
        inline_str_cell(2, 1, "Magnitud", 1), inline_str_cell(2, 2, sensor.get("magnitude", "")),
        inline_str_cell(2, 4, "Tipo paquete", 1), number_cell(2, 5, packet_type),
    ]))
    write_bytes(f, row_xml(3, [
        inline_str_cell(3, 1, "Unidad", 1), inline_str_cell(3, 2, sensor.get("unit", "")),
        inline_str_cell(3, 4, "Flag", 1), number_cell(3, 5, flag),
    ]))
    write_bytes(f, row_xml(4, [
        inline_str_cell(4, 1, "Muestras", 1), number_cell(4, 2, count),
        inline_str_cell(4, 4, "Decodificador", 1), inline_str_cell(4, 5, decoder),
    ]))

    headers = ["tiempo_ms", "tiempo_s"] + [str(x) for x in channel_names]
    write_bytes(f, row_xml(DATA_HEADER_ROW_EXCEL, [inline_str_cell(DATA_HEADER_ROW_EXCEL, i + 1, h, 1) for i, h in enumerate(headers)]))

    row_num = DATA_FIRST_ROW_EXCEL
    buffer: List[str] = []
    flush_every = 5000

    for _, p_type, p_size, body in iter_qul_packets(raw, data_start):
        code, p_flag = body_code_flag(body)
        if p_type != packet_type or code != raw_code or p_flag != flag or p_size != size:
            continue
        t_ms = timestamp_ms_from_body(body)
        values = decode_qul_values(body, decoder)
        cells = [
            f'<c r="A{row_num}"><v>{t_ms}</v></c>',
            f'<c r="B{row_num}"><v>{t_ms / 1000.0:.3f}</v></c>',
        ]
        for idx, value in enumerate(values, start=3):
            if isinstance(value, float) and not math.isfinite(value):
                cells.append(f'<c r="{cell_ref(row_num, idx)}"/>')
            else:
                cells.append(f'<c r="{cell_ref(row_num, idx)}"><v>{format(float(value), ".12g")}</v></c>')
        buffer.append(f'<row r="{row_num}">' + "".join(cells) + '</row>\n')
        row_num += 1
        if len(buffer) >= flush_every:
            write_bytes(f, "".join(buffer))
            buffer.clear()

    if buffer:
        write_bytes(f, "".join(buffer))

    sheet_end(f)


def write_qul_time_sync_sheet_xml(f: BinaryIO, raw: bytes, data_start: int) -> None:
    sheet_start(f, freeze=True, widths=[(1, 1, 14), (2, 3, 18), (4, 4, 22)])
    headers = ["tiempo_ms", "tiempo_s", "unix_s", "offset_s"]
    write_bytes(f, row_xml(1, [inline_str_cell(1, 1, "Sincronización de tiempo .qul", 1)]))
    write_bytes(f, row_xml(DATA_HEADER_ROW_EXCEL, [inline_str_cell(DATA_HEADER_ROW_EXCEL, i + 1, h, 1) for i, h in enumerate(headers)]))
    row_num = DATA_FIRST_ROW_EXCEL
    buffer: List[str] = []
    for _, packet_type, size, body in iter_qul_packets(raw, data_start):
        if packet_type != 2 or len(body) < 15:
            continue
        t_ms = timestamp_ms_from_body(body)
        unix_s = struct.unpack_from("<d", body, 6)[0]
        offset_s = unix_s - (t_ms / 1000.0)
        buffer.append(
            f'<row r="{row_num}">'
            f'<c r="A{row_num}"><v>{t_ms}</v></c>'
            f'<c r="B{row_num}"><v>{t_ms / 1000.0:.3f}</v></c>'
            f'<c r="C{row_num}"><v>{format(unix_s, ".12f")}</v></c>'
            f'<c r="D{row_num}"><v>{format(offset_s, ".12f")}</v></c>'
            f'</row>\n'
        )
        row_num += 1
        if len(buffer) >= 5000:
            write_bytes(f, "".join(buffer))
            buffer.clear()
    if buffer:
        write_bytes(f, "".join(buffer))
    sheet_end(f)


def write_qul_config_sheet_xml(f: BinaryIO, xml_bytes: bytes, tail_text: str) -> None:
    sheet_start(f, freeze=False, widths=[(1, 1, 10), (2, 2, 120)])
    write_bytes(f, row_xml(1, [inline_str_cell(1, 1, "Línea", 1), inline_str_cell(1, 2, "Contenido", 1)]))
    text = xml_bytes.decode("utf-8", errors="replace") + "\n" + tail_text
    for idx, line in enumerate(text.splitlines(), start=2):
        write_bytes(f, row_xml(idx, [number_cell(idx, 1, idx - 1), inline_str_cell(idx, 2, line[:32767])]))
    sheet_end(f)


def build_qul_sheet_plan(packet_infos: List[Dict[str, object]]) -> Tuple[List[str], List[Tuple[str, Optional[Dict[str, object]]]]]:
    used = set()
    sheet_names: List[str] = []
    plan: List[Tuple[str, Optional[Dict[str, object]]]] = []

    sheet_names.append(sanitize_excel_sheet_name("Resumen", used))
    plan.append(("resumen_qul", None))

    for info in packet_infos:
        base = packet_sheet_base_name(info)
        total = int(info["count"])
        parts = max(1, math.ceil(total / ROWS_PER_SHEET))
        if parts == 1:
            sheet_names.append(sanitize_excel_sheet_name(base, used))
            plan.append(("packet_qul", info))
        else:
            # No debería ocurrir en archivos normales, pero se conserva el patrón.
            for part in range(parts):
                info_part = dict(info)
                info_part["part_start"] = part * ROWS_PER_SHEET
                info_part["part_end"] = min((part + 1) * ROWS_PER_SHEET, total)
                sheet_names.append(sanitize_excel_sheet_name(f"{base}_p{part + 1}", used))
                plan.append(("packet_qul", info_part))

    sheet_names.append(sanitize_excel_sheet_name("Sincronizacion_tiempo", used))
    plan.append(("time_qul", None))
    sheet_names.append(sanitize_excel_sheet_name("Configuracion_XML", used))
    plan.append(("config_qul", None))
    return sheet_names, plan


def create_qul_xlsx_native(output_path: Path, qul_path: Path, parsed: Dict[str, object], counts: Dict[Tuple[int, int, int, int], int], packet_infos: List[Dict[str, object]]) -> None:
    raw = parsed["raw"]
    xml_bytes = parsed["xml_bytes"]
    tail_text = parsed["tail_text"]
    time_meta = parsed["time_meta"]
    data_start = int(parsed["data_start"])
    assert isinstance(raw, (bytes, bytearray))
    assert isinstance(xml_bytes, (bytes, bytearray))
    assert isinstance(tail_text, str)
    assert isinstance(time_meta, dict)

    sheet_names, plan = build_qul_sheet_plan(packet_infos)
    sheet_count = len(sheet_names)

    compression_kwargs = {"compression": zipfile.ZIP_DEFLATED, "compresslevel": ZIP_COMPRESSION_LEVEL}
    with zipfile.ZipFile(output_path, "w", allowZip64=True, **compression_kwargs) as z:
        z.writestr("[Content_Types].xml", content_types_xml(sheet_count))
        z.writestr("_rels/.rels", root_rels_xml())
        z.writestr("xl/workbook.xml", workbook_xml(sheet_names))
        z.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml(sheet_count))
        z.writestr("xl/styles.xml", styles_xml())
        z.writestr("docProps/core.xml", core_xml())
        z.writestr("docProps/app.xml", app_xml(sheet_names))

        for idx, (kind, info) in enumerate(plan, start=1):
            sheet_path = f"xl/worksheets/sheet{idx}.xml"
            with z.open(sheet_path, "w", force_zip64=True) as f:
                if kind == "resumen_qul":
                    write_qul_resumen_sheet(f, qul_path, packet_infos, counts, time_meta, data_start)
                elif kind == "packet_qul":
                    assert info is not None
                    write_qul_packet_sheet_xml(f, info, raw, data_start)
                elif kind == "time_qul":
                    write_qul_time_sync_sheet_xml(f, raw, data_start)
                elif kind == "config_qul":
                    write_qul_config_sheet_xml(f, bytes(xml_bytes), tail_text)


def extract_one_qul(qul_path: Path) -> Path:
    """Extrae un archivo .qul y crea su Excel de resultados."""
    output_path = output_folder() / f"{qul_path.stem}_resultados_extraidos.xlsx"

    if output_path.exists():
        try:
            output_path.unlink()
        except PermissionError:
            raise PermissionError(
                f"No puedo sobrescribir '{output_path.name}'. Cierra ese Excel si está abierto y vuelve a ejecutar."
            )

    print(f"\nProcesando: {qul_path.name}")
    print(f"Salida:     {output_path.name}")
    parsed = parse_qul_container(qul_path)
    raw = parsed["raw"]
    data_start = int(parsed["data_start"])
    sensors = parsed["sensors"]
    assert isinstance(raw, (bytes, bytearray))
    assert isinstance(sensors, dict)

    counts = collect_qul_counts(bytes(raw), data_start)
    packet_infos = build_qul_packet_infos(counts, sensors)
    if not packet_infos:
        raise RuntimeError("No encontré paquetes de datos exportables dentro del archivo .qul.")

    for info in packet_infos:
        sensor = info["sensor"]
        assert isinstance(sensor, dict)
        print(f"  - Tipo {info['packet_type']} / código {info['raw_code']}: {sensor.get('name', '')} ({info['count']} muestras)")

    print("  Creando Excel...")
    create_qul_xlsx_native(output_path, qul_path, parsed, counts, packet_infos)
    print(f"OK: {output_path}")
    return output_path


def find_wimu_input_files(wimu_folder: Path) -> List[Path]:
    files: List[Path] = []
    for ext in QUL_EXTENSIONS:
        files.extend(wimu_folder.glob(f"*{ext}"))
        files.extend(wimu_folder.glob(f"*{ext.upper()}"))
    # Elimina duplicados manteniendo orden: .qul primero, luego .qui.
    unique: List[Path] = []
    seen = set()
    for p in sorted(files, key=lambda x: (x.suffix.lower() != ".qul", x.name.lower())):
        key = str(p.resolve()).lower()
        if key not in seen:
            unique.append(p)
            seen.add(key)
    return unique


def extract_one_file(path: Path) -> Path:
    suffix = path.suffix.lower()
    if suffix == ".qul":
        return extract_one_qul(path)
    if suffix == ".qui":
        return extract_one_qui(path)
    raise RuntimeError(f"Extensión no soportada: {path.suffix}")


def main() -> int:
    try:
        wimu_folder = find_wimu_folder()
        input_files = find_wimu_input_files(wimu_folder)

        if not input_files:
            print(f"ERROR: no encontré archivos .qul ni .qui en: {wimu_folder}")
            return 1

        print("\n" + "=" * 60)
        print("SportsLabResearch-WIMU-DataExtractor v1.0.0")
        print("=" * 60)

        print("\nArchivos encontrados:\n")
        for i, f in enumerate(input_files, start=1):
            print(f"{i}. {f.name}")

        print("\n1. Procesar un archivo")
        print("2. Procesar varios archivos")
        print("3. Procesar todos los archivos")
        print("0. Salir")

        opcion = input("\nSeleccione una opción: ").strip()

        if opcion == "0":
            return 0
        elif opcion == "1":
            seleccion = input("Número del archivo: ").strip()
            archivos = [input_files[int(seleccion)-1]]
        elif opcion == "2":
            seleccion = input("Números separados por comas (ej. 1,3,5): ").strip()
            indices = [int(x.strip())-1 for x in seleccion.split(",")]
            archivos = [input_files[i] for i in indices]
        elif opcion == "3":
            archivos = input_files
        else:
            print("Opción no válida.")
            return 1

        outputs = []
        for archivo in archivos:
            outputs.append(extract_one_file(archivo))

        print("\nArchivos generados:")
        for out in outputs:
            print(f" - {out}")

        return 0

    except Exception as exc:
        print(f"\nERROR: {exc}")
        return 1

if __name__ == "__main__":
    raise SystemExit(main())
