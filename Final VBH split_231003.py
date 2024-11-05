import os
from PyPDF2 import PdfFileMerger
import pdfplumber

dir_list = ['230602 - Final Logs', '230710 - Final Logs']

for dir in dir_list:
    for i in range (0, 3):
        info = ["logs", "labs", "photos"]
        contractor = ["AHAM", "ERI"]
        #please select info
        info_choice = info[i]
        contractor_choice = contractor[1]
        finalLogs_dir = f"C:/Users/921722/Box/BI3753 Team/30 - Geotech/06-Grouting works/02 - Verification/01 Copy of all verification BHs/01 Final Logs to Check/{dir}"
        finalLogs_list = os.listdir(finalLogs_dir)
        print(finalLogs_list)
        if "split" in finalLogs_list:
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

                print(dir, info_choice[i], textPage_width_list)


                if info_choice == "logs":
                    merger.write(f"{finalLogs_dir}\split\{info_choice}\{VBH_name}.pdf")

                if info_choice == "labs":
                    merger.write(f"{finalLogs_dir}\split\{info_choice}\{VBH_name}.pdf")

                if info_choice == "photos":
                    merger.write(f"{finalLogs_dir}\split\{info_choice}\{VBH_name}.pdf")

                merger.close()

        else:

            for folder in finalLogs_list:
                finalLogs_dir_sub = f"C:/Users/921722/Box/BI3753 Team/30 - Geotech/06-Grouting works/02 - Verification/01 Copy of all verification BHs/01 Final Logs to Check/{dir}/{folder}"
                print(finalLogs_dir_sub)
                if not os.path.isdir(finalLogs_dir_sub):
                    continue
                finalLogs_dir_sub_list = os.listdir(finalLogs_dir_sub)
                for finalLogs_full in finalLogs_dir_sub_list:
                    merger = PdfFileMerger()
                    VBH_name = finalLogs_full.replace(".pdf", "")
                    if not finalLogs_full.endswith(".pdf"):
                        continue
                    finalLogs_file_path = f'{finalLogs_dir_sub}\{finalLogs_full}'
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
                                        merger.append(open(finalLogs_file_path, 'rb'), pages=(k, k + 1))

                                if info_choice == "labs":
                                    if textPage.width == 792 or textPage.width == 841.68:
                                        merger.append(open(finalLogs_file_path, 'rb'), pages=(k, k + 1))

                                if info_choice == "photos":
                                    if textPage.width == 780:
                                        merger.append(open(finalLogs_file_path, 'rb'), pages=(k, k + 1))

                            if contractor_choice == "ERI":
                                if textPage.width == 841.68:
                                    ERI_labs_flip = "Y"
                                if info_choice == "logs":
                                    if textPage.width == 612:
                                        merger.append(open(finalLogs_file_path, 'rb'), pages=(k, k + 1))

                                if info_choice == "labs":
                                    if (
                                            textPage.width == 595.439994 and ERI_labs_flip == "Y") or textPage.width == 841.68:
                                        merger.append(open(finalLogs_file_path, 'rb'), pages=(k, k + 1))

                                if info_choice == "photos":
                                    if textPage.width == 595.439994 and ERI_labs_flip == "N":
                                        merger.append(open(finalLogs_file_path, 'rb'), pages=(k, k + 1))

                    print(finalLogs_full, textPage_width_list)

                    if info_choice == "logs":
                        merger.write(f"{finalLogs_dir_sub}\split\{info_choice}\{VBH_name}.pdf")

                    if info_choice == "labs":
                        merger.write(f"{finalLogs_dir_sub}\split\{info_choice}\{VBH_name}.pdf")

                    if info_choice == "photos":
                        merger.write(f"{finalLogs_dir_sub}\split\{info_choice}\{VBH_name}.pdf")

                    merger.close()