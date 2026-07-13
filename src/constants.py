# -*- coding: utf-8 -*-

"""
Constantes globales de SportsLabResearch-WIMU-DataExtractor.
"""

from __future__ import annotations

EXCEL_MAX_ROWS = 1_048_576
DATA_HEADER_ROW_EXCEL = 7
DATA_FIRST_ROW_EXCEL = 8
ROWS_PER_SHEET = EXCEL_MAX_ROWS - DATA_FIRST_ROW_EXCEL + 1

INCLUDE_TIME_SECONDS = True
ZIP_COMPRESSION_LEVEL = 1

SUPPORTED_EXTENSIONS = (".qul", ".qui")

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PACKAGE_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"

QUL_MAGIC = b"\xe1\xed"

QUL_ALIAS_SENSOR_CODE = {
    44: 300,
    45: 301,
    46: 302,
}

QUL_ORDER = {
    (1, 81): 10,
    (6, 44): 20,
    (6, 46): 30,
    (6, 45): 40,
    (6, 25): 50,
    (6, 31): 60,
    (6, 62): 70,
    (6, 41): 80,
    (6, 42): 90,
}

APP_NAME = "SportsLabResearch-WIMU-DataExtractor"
VERSION = "1.0.0-alpha"
LICENSE = "MIT"
AUTHOR = "SportsLabResearch"