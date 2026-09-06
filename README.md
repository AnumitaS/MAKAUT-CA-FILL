# MAKAUT-CA-FILL 

## Overview
This project helps you automate the creation of MAKAUT Continuous Assessment Mark Submission PDF files. You need to provide an Excel (`.xlsx`) or ODS (`.ods`) sheet containing obtained marks, names, and roll numbers of students using the exact column names supplied with this project.

## How to Run
1. Run `SCRIPT.py`:
   ```bash
   python SCRIPT.py
   ```

2. You will be prompted to enter the following information manually:
   * **Academic Year** *(e.g., 2026-27)*
   * **Semester** *(e.g., 5th)*
   * **Programme** *(e.g., B.Tech., ECE)*
   * **Subject Name** *(e.g., Web Technology)*
   * **Paper Code** *(e.g., OE-EC704A)*
   * **UPID** *(e.g., 007716)*
   * **Exam Date** *(e.g., 02/09/26)*
   * **Subject Teacher Name**
   * **Teacher Mobile Number**

3. Next, select the required resource files via the file picker prompts:
   * `template.docx` file
   * Teacher signature image file
   * College stamp image file
   * Student data file (`.xlsx` or `.ods`)

> **Note:** The column names in your student data file must match the sample file provided in this repository (`subject_name/input/data_files.xlsx`), as they are tagged directly inside `SCRIPT.py`. You must select the correct files; otherwise, execution will fail.

## Output Directory
Generated files will be saved in the directory structure:
`root(MAKAUT_CA)/subject(CSE)/out/`
