# -*- coding: utf-8 -*-

"""
SportsLabResearch-WIMU-DataExtractor

Open-source extractor for WIMU .QUL and .QUI files.
"""

__version__ = "1.0.0-alpha"
__author__ = "SportsLabResearch"
__license__ = "MIT"

from .extractor import WIMUExtractor

__all__ = [
    "WIMUExtractor",
]