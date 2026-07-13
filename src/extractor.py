# -*- coding: utf-8 -*-

"""
SportsLabResearch-WIMU-DataExtractor

Core extractor module.
Version: v1.0.0-alpha
"""

from pathlib import Path

SUPPORTED_EXTENSIONS = (".qul", ".qui")


class WIMUExtractor:
    """Core class for WIMU file discovery."""

    def __init__(self, root: Path):
        self.root = Path(root)

    def scan(self) -> list[Path]:
        files: list[Path] = []

        for ext in SUPPORTED_EXTENSIONS:
            files.extend(self.root.rglob(f"*{ext}"))
            files.extend(self.root.rglob(f"*{ext.upper()}"))

        return sorted(set(files))

    @staticmethod
    def is_supported(path: Path) -> bool:
        return path.suffix.lower() in SUPPORTED_EXTENSIONS

    @staticmethod
    def print_summary(files: list[Path]) -> None:
        print()
        print("=" * 72)
        print("DISCOVERED WIMU FILES")
        print("=" * 72)

        if not files:
            print("No compatible files found.")
            return

        for i, file in enumerate(files, start=1):
            print(f"{i:3d}. {file.name}")

        print("-" * 72)
        print(f"Total files : {len(files)}")
        print("=" * 72)