import shutil
import os
from os import path
from PyPDF2 import PdfFileMerger
import pdfplumber

# landsideDrawings_dir = "C:\\Users\\921722\\Box\\BI3753 Team\\61 - Received Info ADAC-CMA\\220225 Landside as built drawings\\Landside Utility As Built Drqawing PDF and CAD"
# landsideDrawings_list = os.listdir(landsideDrawings_dir)
# paste_path_0 = "C:\\Users\\921722\\OneDrive - Royal HaskoningDHV\\20221024 AD\\Landside\\ITD"
#
# for landsideDrawing_files in landsideDrawings_list:
#     if landsideDrawing_files.endswith(".pdf"):
#         continue
#     if landsideDrawing_files.endswith(".pdf"):
#         continue
#     if landsideDrawing_files.endswith(".dwg"):
#         continue
#     if landsideDrawing_files.endswith(".xlsx"):
#         continue
#     name_string = landsideDrawing_files.split("-")
#
#     if name_string[6] == "ITEX":
#         drawing_path = f"{landsideDrawings_dir}\{landsideDrawing_files}\{landsideDrawing_files}.pdf"
#         paste_path = f"{paste_path_0}\{landsideDrawing_files}.pdf"
#         print(path.exists(drawing_path))
#         shutil.copyfile(drawing_path, paste_path)
dir = f"C:\\Users\\921722\\OneDrive - Royal HaskoningDHV\\20221024 AD\\Landside\PWD"
list = os.listdir(dir)
print(list)

merger = PdfFileMerger()

for file in list:
    file_path = f"{dir}\{file}"
    with pdfplumber.open(file_path) as pdf:
        totalPages = len(pdf.pages)
        for k in range(totalPages):
            merger.append(open(file_path, 'rb'), pages=(k, k+1))

merger.write(f"C:\\Users\\921722\\OneDrive - Royal HaskoningDHV\\20221024 AD\\Landside\\PWD\\merged.pdf")