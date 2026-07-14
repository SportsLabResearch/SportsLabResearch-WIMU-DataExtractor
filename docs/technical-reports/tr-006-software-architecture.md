# Technical Report TR-006 — Software Architecture

Describe the architecture of SportsLabResearch-WIMU-DataExtractor.

## Purpose

Provide an overview of the software structure and the interaction between its main modules.

## Architecture

    WIMU® Data
         │
         ▼
    Data Import
         │
         ▼
    Extraction Engine
         │
         ▼
    Processing Engine
         │
         ▼
    Validation Engine
         │
         ▼
    Reporting Engine
         │
         ▼
    Excel / Word Outputs

## Software Modules

| Module | Function |
|--------|----------|
| Data Import | Loads WIMU® datasets. |
| Extraction Engine | Reads supported variables. |
| Processing Engine | Standardizes datasets and variables. |
| Validation Engine | Compares outputs with reference datasets. |
| Reporting Engine | Generates Excel and Word reports. |

## Design Principles

- Modular architecture.
- Independent processing stages.
- Standardized outputs.
- Scientific reproducibility.
- Easy maintenance and scalability.

## Related Documentation

- TR-002 Processing Workflow
- TR-003 Variable Dictionary
- TR-004 Validation Protocol
- TR-005 Output Specification
