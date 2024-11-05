import os
from PyPDF2 import PdfFileMerger
import pdfplumber


info = ["logs", "labs", "photos"]
contractor = ["AHAM", "ERI"]
#please select info
info_choice = info[2]
contractor_choice = contractor[0]
#######
finalLogs_dir = f"C:/Users/921722/Box/BI3753 Team/30 - Geotech/06-Grouting works/02 - Verification/01 Copy of all verification BHs/01 Final Logs to Check/230209/{contractor_choice}"
finalLogs_list = os.listdir(finalLogs_dir)
print(finalLogs_list)

for finalLogs_full in finalLogs_list:
    merger = PdfFileMerger()
    VBH_name = finalLogs_full.replace(".pdf", "")
    if not finalLogs_full.endswith(".pdf"):
        continue
    finalLogs_file_path = f'{finalLogs_dir}\{finalLogs_full}'
    textPage_width_list = []
    with pdfplumber.open(finalLogs_file_path) as pdf:
        totalPages = len(pdf.pages)
        ERI_labs_flip = "N"
        for k in range(totalPages):
            textPage = pdf.pages[k]
            textPage_width_list.append(textPage.width)

            if contractor_choice == "AHAM":
                if info_choice == "logs":
                    if textPage.width == 595.2:
                        merger.append(open(finalLogs_file_path, 'rb'), pages=(k, k+1))

                if info_choice == "labs":
                    if textPage.width == 792 or textPage.width == 841.68:
                        merger.append(open(finalLogs_file_path, 'rb'), pages=(k, k+1))

                if info_choice == "photos":
                    if textPage.width == 780:
                        merger.append(open(finalLogs_file_path, 'rb'), pages=(k, k+1))

            if contractor_choice == "ERI":
                if textPage.width == 841.68:
                    ERI_labs_flip = "Y"
                if info_choice == "logs":
                    if textPage.width == 612:
                        merger.append(open(finalLogs_file_path, 'rb'), pages=(k, k+1))

                if info_choice == "labs":
                    if (textPage.width == 595.439994 and ERI_labs_flip == "Y") or textPage.width == 841.68:
                        merger.append(open(finalLogs_file_path, 'rb'), pages=(k, k+1))

                if info_choice == "photos":
                    if textPage.width == 595.439994 and ERI_labs_flip == "N":
                        merger.append(open(finalLogs_file_path, 'rb'), pages=(k, k+1))

    print(finalLogs_full, textPage_width_list)


    if info_choice == "logs":
        merger.write(f"C:/Users/921722/Box/BI3753 Team/30 - Geotech/06-Grouting works/02 - Verification/01 Copy of all verification BHs/01 Final Logs to Check/230209/{contractor_choice}/{info_choice}/{VBH_name}.pdf")

    if info_choice == "labs":
        merger.write(f"C:/Users/921722/Box/BI3753 Team/30 - Geotech/06-Grouting works/02 - Verification/01 Copy of all verification BHs/01 Final Logs to Check/230209/{contractor_choice}/{info_choice}/{VBH_name}.pdf")

    if info_choice == "photos":
        merger.write(f"C:/Users/921722/Box/BI3753 Team/30 - Geotech/06-Grouting works/02 - Verification/01 Copy of all verification BHs/01 Final Logs to Check/230209/{contractor_choice}/{info_choice}/{VBH_name}.pdf")

    merger.close()
