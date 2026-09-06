import os
import io
import platform
import subprocess
import sys
from pathlib import Path
import pandas as pd
import openpyxl
from openpyxl_image_loader import SheetImageLoader
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches

# Native GUI File Selection Prompts
import tkinter as tk
from tkinter import filedialog

# Hide the main Tkinter root window
root = tk.Tk()
root.withdraw()
root.attributes('-topmost', True)

# Detect Operating System
IS_WINDOWS = platform.system() == "Windows"

if IS_WINDOWS:
    import comtypes.client

print("==================================================")
print("     DOCUMENT & PDF AUTOMATION GENERATOR          ")
print("==================================================\n")

# Helper function to sanitize string values for directory/file naming
def clean_filename_str(val):
    if pd.isna(val) or str(val).strip().lower() in ["none", "nan"]:
        return ""
    text = str(val).strip()
    return text.replace(" ", "_").replace("/", "_").replace("\\", "_").replace(".", "")

# 1. Interactive CLI Inputs for Common Metadata
print("--- Step 1: Enter Subject Metadata ---")
common_data = {
    "Year": input("Enter Academic Year (e.g., 2026-27): ").strip(),
    "Semester": input("Enter Semester (e.g., 5th): ").strip(),
    "Programme": input("Enter Programme (e.g., B.Tech., ECE): ").strip(),
    "Subject": input("Enter Subject Name (e.g., Web Technology): ").strip(),
    "PaperCode": input("Enter Paper Code (e.g., OE-EC704A): ").strip(),
    "UPID": input("Enter UPID (e.g., 007716): ").strip(),
    "ExamDate": input("Enter Exam Date (e.g., 02/09/26): ").strip(),
    "SubjectTeacher": input("Enter Subject Teacher Name: ").strip(),
    "TeacherMobile": input("Enter Teacher Mobile Number: ").strip()
}

# 2. Graphical File Selection Dialog Prompts
print("\n--- Step 2: Select Files via File Explorer Prompts ---")

print("Opening prompt to select Word Template (.docx)...")
TEMPLATE_PATH = filedialog.askopenfilename(
    title="Select Word Template File",
    filetypes=[("Word Documents", "*.docx")]
)

if not TEMPLATE_PATH:
    print("No template selected. Exiting script.")
    sys.exit()

print("Opening prompt to select Input Data File (.xlsx, .ods, .csv)...")
DATA_FILE_PATH = filedialog.askopenfilename(
    title="Select Input Data File",
    filetypes=[("Data Files", "*.xlsx *.ods *.csv *.xls")]
)

if not DATA_FILE_PATH:
    print("No data file selected. Exiting script.")
    sys.exit()

# Optional Images Selection
print("Opening prompt to select Teacher Signature Image (Optional - Cancel to skip)...")
TEACHER_SIG_PATH = filedialog.askopenfilename(
    title="Select Teacher Signature (Optional)",
    filetypes=[("Image Files", "*.png *.jpg *.jpeg")]
)

print("Opening prompt to select College Stamp Image (Optional - Cancel to skip)...")
COLLEGE_STAMP_PATH = filedialog.askopenfilename(
    title="Select College Stamp (Optional)",
    filetypes=[("Image Files", "*.png *.jpg *.jpeg")]
)

# 3. Dynamic Output Directory Creation: MAKAUT_CA/OUTPUT/FOLDER_NAME/DOC
input_file_path = Path(DATA_FILE_PATH).resolve()

# Find project root directory named 'MAKAUT_CA'
root_dir = None
for parent in [input_file_path] + list(input_file_path.parents):
    if parent.name.upper() == "MAKAUT_CA":
        root_dir = parent
        break

if not root_dir:
    root_dir = input_file_path.parents[2]  # Fallback: 3 levels up from input file

prog_clean = clean_filename_str(common_data["Programme"]) or "Programme"
subj_clean = clean_filename_str(common_data["Subject"]) or "Subject"
year_clean = clean_filename_str(common_data["Year"]) or "Year"

folder_name = f"{prog_clean}_{subj_clean}_{year_clean}"

# Target Output Paths: MAKAUT_CA/OUTPUT/FOLDER_NAME/DOC and MAKAUT_CA/OUTPUT/FOLDER_NAME/PDF
OUTPUT_DIR_DOCX = os.path.join(root_dir, "OUTPUT", folder_name, "DOC")
OUTPUT_DIR_PDF = os.path.join(root_dir, "OUTPUT", folder_name, "PDF")

for directory in [OUTPUT_DIR_DOCX, OUTPUT_DIR_PDF]:
    os.makedirs(directory, exist_ok=True)

print(f"\n[INFO] Selected Template:    {TEMPLATE_PATH}")
print(f"[INFO] Selected Input Data:  {DATA_FILE_PATH}")
print(f"[INFO] Root Directory:       {root_dir}")
print(f"[INFO] Output DOCX Directory: {OUTPUT_DIR_DOCX}")
print(f"[INFO] Output PDF Directory:  {OUTPUT_DIR_PDF}")

# Helper Data Formatting Functions
def clean_val(val):
    if pd.isna(val) or str(val).strip().lower() in ["none", "nan"]:
        return ""
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    return str(val).strip()

def clean_mark(val):
    cleaned = clean_val(val)
    return cleaned if cleaned != "" else " "

def normalize_col(name):
    """Normalizes column names by stripping spaces, underscores, and lowercasing."""
    return str(name).lower().replace("_", "").replace(" ", "").strip()

def get_by_flexible_col_name(row, target_names, default=""):
    if isinstance(target_names, str):
        target_names = [target_names]
        
    normalized_targets = [normalize_col(t) for t in target_names]
    
    for col in row.index:
        if normalize_col(col) in normalized_targets:
            val = row[col]
            return val if not pd.isna(val) else default
    return default

def load_image_by_col_name(file_path, row_idx, target_names, df_columns):
    if isinstance(target_names, str):
        target_names = [target_names]
    normalized_targets = [normalize_col(t) for t in target_names]
    
    matched_col_idx = None
    for idx, col in enumerate(df_columns):
        if normalize_col(col) in normalized_targets:
            matched_col_idx = idx + 1
            break

    if not matched_col_idx:
        return None

    ext = os.path.splitext(file_path)[1].lower()
    
    if ext == ".xlsx":
        try:
            wb = openpyxl.load_workbook(file_path, data_only=True)
            sheet = wb.active
            image_loader = SheetImageLoader(sheet)
            col_letter = openpyxl.utils.get_column_letter(matched_col_idx)
            cell_ref = f"{col_letter}{row_idx}"
            if image_loader.image_in(cell_ref):
                img = image_loader.get(cell_ref)
                img_byte_arr = io.BytesIO()
                img.save(img_byte_arr, format='PNG')
                img_byte_arr.seek(0)
                return img_byte_arr
        except Exception as e:
            print(f"Warning: Could not read XLSX image at row {row_idx}: {e}")
            
    elif ext == ".ods":
        try:
            import zipfile
            with zipfile.ZipFile(file_path, 'r') as z:
                media_files = [f for f in z.namelist() if f.startswith('Pictures/')]
                if row_idx - 2 < len(media_files):
                    return io.BytesIO(z.read(media_files[row_idx - 2]))
        except Exception as e:
            print(f"Warning: Could not read ODS image at row {row_idx}: {e}")
            
    return None

def convert_docx_to_pdf(docx_path, output_pdf_dir, word_app=None):
    docx_path_abs = os.path.abspath(docx_path)
    output_pdf_dir_abs = os.path.abspath(output_pdf_dir)
    file_name = os.path.splitext(os.path.basename(docx_path))[0]
    pdf_path_abs = os.path.join(output_pdf_dir_abs, f"{file_name}.pdf")

    if IS_WINDOWS and word_app:
        try:
            doc = word_app.Documents.Open(docx_path_abs)
            doc.SaveAs(pdf_path_abs, FileFormat=17)
            doc.Close()
        except Exception as e:
            print(f"Error converting {docx_path} on Windows: {e}")
    else:
        try:
            subprocess.run(
                ["libreoffice", "--headless", "--convert-to", "pdf", docx_path_abs, "--outdir", output_pdf_dir_abs],
                check=True,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL
            )
        except Exception as e:
            print(f"Error converting {docx_path} on Linux: {e}")

# 4. Read Input Data File (.xlsx, .ods, .csv)
ext = os.path.splitext(DATA_FILE_PATH)[1].lower()
if ext == ".csv":
    df = pd.read_csv(DATA_FILE_PATH)
elif ext in [".xlsx", ".xls"]:
    df = pd.read_excel(DATA_FILE_PATH)
elif ext == ".ods":
    df = pd.read_excel(DATA_FILE_PATH, engine="odf")
else:
    raise ValueError(f"Unsupported file format: {ext}")

df.columns = df.columns.astype(str).str.strip()

print(f"\n[INFO] Operating System: {platform.system()}")
print(f"[INFO] Columns Detected in Excel: {list(df.columns)}")
print(f"[INFO] Total Rows Detected: {len(df)}\n")

word_app = None
if IS_WINDOWS:
    word_app = comtypes.client.CreateObject('Word.Application')
    word_app.Visible = False

processed_count = 0

try:
    # 5. Process Rows & Generate Files
    for idx, row in df.iterrows():
        row_idx = idx + 2  # Row index in Excel sheet (1-based header = 1)

        student_name = get_by_flexible_col_name(row, ["Student name", "StudentName", "Student Name", "Name"])
        roll_number = get_by_flexible_col_name(row, ["Roll", "RollNumber", "Roll Number", "RollNo", "Roll_No"])

        raw_name = clean_val(student_name)
        raw_roll = clean_val(roll_number)

        if not raw_name and not raw_roll:
            print(f"[SKIP] Row {row_idx}: Empty student record.")
            continue

        doc = DocxTemplate(TEMPLATE_PATH)
        context = common_data.copy()

        context.update({
            "StudentName": raw_name,
            "RollNumber": raw_roll,
            "q_i": clean_mark(get_by_flexible_col_name(row, ["q_i", "qi", "q1i"])),
            "q_ii": clean_mark(get_by_flexible_col_name(row, ["q_ii", "qii", "q1ii"])),
            "q_iii": clean_mark(get_by_flexible_col_name(row, ["q_iii", "qiii", "q1iii"])),
            "q_iv": clean_mark(get_by_flexible_col_name(row, ["q_iv", "qiv", "q1iv"])),
            "q_v": clean_mark(get_by_flexible_col_name(row, ["q_v", "qv", "q1v"])),
            "q_2": clean_mark(get_by_flexible_col_name(row, ["q_2", "q2"])),
            "q_3": clean_mark(get_by_flexible_col_name(row, ["q_3", "q3"])),
            "q_4": clean_mark(get_by_flexible_col_name(row, ["q_4", "q4"])),
            "q_5": clean_mark(get_by_flexible_col_name(row, ["q_5", "q5"])),
            "q_6": clean_mark(get_by_flexible_col_name(row, ["q_6", "q6"])),
            "q_7": clean_mark(get_by_flexible_col_name(row, ["q_7", "q7"])),
            "TotalMarks": clean_mark(get_by_flexible_col_name(row, ["TotalMarks", "Total Marks", "Total"]))
        })

        if TEACHER_SIG_PATH and os.path.exists(TEACHER_SIG_PATH):
            context["TeacherSig"] = InlineImage(doc, TEACHER_SIG_PATH, width=Inches(1.2))
        else:
            context["TeacherSig"] = ""

        if COLLEGE_STAMP_PATH and os.path.exists(COLLEGE_STAMP_PATH):
            context["CollegeStamp"] = InlineImage(doc, COLLEGE_STAMP_PATH, width=Inches(1.5))
        else:
            context["CollegeStamp"] = ""

        img_bytes = load_image_by_col_name(DATA_FILE_PATH, row_idx, ["Student sign", "StudentSig", "Student Sig", "Signature"], df.columns)
        if img_bytes:
            context["StudentSig"] = InlineImage(doc, img_bytes, width=Inches(1.2))
        else:
            context["StudentSig"] = ""

        doc.render(context)
        
        # Output filename format: Name_Roll_Subject
        clean_name = clean_filename_str(raw_name) or f"Student_{row_idx}"
        clean_roll = clean_filename_str(raw_roll) or str(row_idx)
        
        output_base_name = f"{clean_name}_{clean_roll}_{subj_clean}"
        
        # Save DOCX
        docx_filename = os.path.join(OUTPUT_DIR_DOCX, f"{output_base_name}.docx")
        doc.save(docx_filename)
        
        # Convert & Save PDF
        convert_docx_to_pdf(docx_filename, OUTPUT_DIR_PDF, word_app)
        print(f"[SUCCESS] Row {row_idx}: Generated DOCX & PDF -> {output_base_name}")
        processed_count += 1

finally:
    if IS_WINDOWS and word_app:
        word_app.Quit()

print(f"\nProcessing complete! {processed_count} student records generated successfully.")
