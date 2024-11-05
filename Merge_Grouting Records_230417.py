import os
from PyPDF2 import PdfFileMerger

gr_dir = r"C:\Users\921722\Box\BI3753 Team\30 - Geotech\06-Grouting works\01 - Grouting Records\Pier 2\Pier 2 Grouting Records Received 19-04-2022\34) NP-12"
gr_utils = os.listdir(gr_dir)
merger = PdfFileMerger()

for file in gr_utils:
    if "--.pdf" not in file:
        continue
    individual_path = f"{gr_dir}\{file}"
    print(file, file)
    merger.append(open(individual_path, 'rb'))

merger.write(f"{gr_dir}\\NP12_Merged Permeation Record.pdf")