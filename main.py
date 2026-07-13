# -*- coding: utf-8 -*-

"""
SportsLabResearch-WIMU-DataExtractor
v1.0.0-alpha
"""

from pathlib import Path

from src.extractor import WIMUExtractor

APP_NAME = "SportsLabResearch-WIMU-DataExtractor"
VERSION = "v1.0.0-alpha"


def banner() -> None:
    print("=" * 72)
    print(APP_NAME)
    print(VERSION)
    print("=" * 72)
    print()


def main() -> None:
    banner()

    project = Path(__file__).resolve().parent
    data_input = project / "data" / "input"

    extractor = WIMUExtractor(data_input)

    files = extractor.scan()

    extractor.print_summary(files)


if __name__ == "__main__":
    main()