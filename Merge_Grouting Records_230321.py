import os
from PyPDF2 import PdfFileMerger

gr_dir = r"C:\Users\921722\Box\BI3753 Team\30 - Geotech\06-Grouting works\01 - Grouting Records\Pier 3\PIRE-3_Grouting_Record"
gr_utils = os.listdir(gr_dir)

for util in gr_utils:
    merger = PdfFileMerger()
    if "3_" not in util:
        continue
    gr_folder_path = f"{gr_dir}\{util}"
    gr_folders = os.listdir(gr_folder_path)
    for folder in gr_folders:
        gr_file_path = f"{gr_folder_path}\{folder}"
        gr_files = os.listdir(gr_file_path)
        for file in gr_files:
            if ".pdf" not in file:
                continue
            individual_path = f"{gr_file_path}\{file}"
            print(util, file)
            merger.append(open(individual_path, 'rb'))
    merger.write(f"{gr_dir}\Merged PDF\{util}.pdf")