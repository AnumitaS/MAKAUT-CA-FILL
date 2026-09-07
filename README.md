# MAKAUT-CA-FILL
## Tutorial video
* **https://www.youtube.com/watch?v=lQmGZlvivZw**
* [![Watch the video](https://img.youtube.com/vi/lQmGZlvivZw/maxresdefault.jpg)](https://www.youtube.com/watch?v=lQmGZlvivZw)

## Overview
This project automates the generation of MAKAUT Continuous Assessment (CA) Mark Submission documents (DOCX) and PDF files. You need to provide an Excel (`.xlsx`), ODS (`.ods`), or CSV (`.csv`) sheet containing student marks, names, roll numbers, and optional signatures using column names compatible with the project template.

---

## Output Directory Structure
The script automatically detects the root `MAKAUT_CA` directory and saves all generated documents into the following folder structure:

```text
MAKAUT_CA/
└── requirements.txt
└── College_Stamp_image.png
└── Teacher_sign_image.png
└── MAKAUT_Template.docx
└── OUTPUT/
    └── <PROGRAMME>_<SUBJECT>_<YEAR>/
        ├── DOC/
        │   ├── StudentName_RollNumber_Subject.docx
        │   └── ...
        └── PDF/
            ├── StudentName_RollNumber_Subject.pdf
            └── ...
```

*Example Output Folder Path:*  
`MAKAUT_CA/OUTPUT/BTech_ECE_Web_Technology_2026-27/DOC`  
`MAKAUT_CA/OUTPUT/BTech_ECE_Web_Technology_2026-27/PDF`

---

## How to Run

1. **Install required libraries**
    * ***pandas>=2.0.0***
    * ***openpyxl>=3.1.0***
    * ***openpyxl-image-loader>=1.0.5***
    * ***docxtpl>=0.16.0***
    * ***python-docx>=0.8.11***
    * ***comtypes>=1.2.0***
    * ***odfpy>=1.4.1***

3. **Execute the Script:**
   ```bash
   python SCRIPT.py
   ```

4. **Enter Subject Metadata via Terminal Prompts:**
   * **Academic Year** *(e.g., 2026-27)*
   * **Semester** *(e.g., 5th)*
   * **Programme** *(e.g., B.Tech., ECE)*
   * **Subject Name** *(e.g., Web Technology)*
   * **Paper Code** *(e.g., OE-EC704A)*
   * **UPID** *(e.g., 007716)*
   * **Exam Date** *(e.g., 02/09/26)*
   * **Subject Teacher Name**
   * **Teacher Mobile Number**

5. **Select Resource Files via GUI File Pickers:**
   * **Word Template:** Select your `.docx` template file.
   * **Student Data File:** Select your `.xlsx`, `.ods`, or `.csv` data sheet.
   * **Teacher Signature Image:** *(Optional - click Cancel to skip)*
   * **College Stamp Image:** *(Optional - click Cancel to skip)*

---

## Requirements & Dependencies

Install required Python packages before running the script:

```bash
pip install pandas openpyxl openpyxl-image-loader docxtpl python-docx comtypes odfpy
```

* **Windows OS:** Uses Microsoft Word (`comtypes`) for native DOCX to PDF conversion.
* **Linux/macOS:** Requires **LibreOffice** installed for headless PDF conversion (`libreoffice --headless`).
