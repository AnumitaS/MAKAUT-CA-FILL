# MAKAUT-CA-FILL
This project will help you create MAKAUT Continuous Assessment Mark Submission pdf file creation. You need to provide a excel, ods sheet with obtained marks, name, roll number of student in exact given column namme supplied with this project

How to run:
Run the SCRIPT.py
You will be prompt to put following information manually:

  Enter Academic Year (e.g., 2026-27): 
  Enter Semester (e.g., 5th): 
  Enter Programme (e.g., B.Tech., ECE): 
  Enter Subject Name (e.g., Web Technology): 
  Enter Paper Code (e.g., OE-EC704A): 
  Enter UPID (e.g., 007716): 
  Enter Exam Date (e.g., 02/09/26): 
  Enter Subject Teacher Name: 
  Enter Teacher Mobile Number: 

After that you will be asked to select following resource files:
  templafe.docx file
  teacher signature image file 
  college stamp image file
  student data file (excel, ods, remmebr the column name must like one given with this repository as the column names are tagged in the script.py , is subject_name/input folder/ data_files.xlsx )
You mus select correct files, else it will fail
Your files will be generated inside root(MAKAUT_CA)/subject(CSE)/output_folder
