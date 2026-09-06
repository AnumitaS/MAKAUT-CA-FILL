# MAKAUT Continuous Assessment Document & PDF Generator

An automated, cross-platform tool designed to simplify the generation of Continuous Assessment (CA) mark submission documents for MAKAUT. It processes student mark sheets, renders populated Word templates (`.docx`), inserts relevant signatures or stamps, and converts the generated files into production-ready PDFs.

---

## Features

* **Cross-Platform Compatibility:** Runs on both **Windows** (using Microsoft Word) and **Ubuntu/Linux** (using LibreOffice headless rendering).
* **Multi-Format Data Ingestion:** Supports `.xlsx`, `.ods`, and `.csv` input datasets.
* **Interactive Native File Chooser:** File pickers allow you to select templates, mark sheets, signatures, and stamps.
* **Automated Output Structuring:** Output files are created outside the input folder at `../out/doc/` and `../out/pdf/`.
* **Standardized File Naming:** Output files follow the `[StudentName]_[RollNumber]_[Subject].docx` (and `.pdf`) pattern.

---

## Prerequisites

### 1. Python Environment
Ensure you have Python 3.8+ installed along with the required dependencies:

```bash
pip install pandas openpyxl openpyxl-image-loader docxtpl python-docx pyexcel-ods3 comtypes
