import os
import pandas as pd

GR_saved_path = r'C:\Users\921722\Box\BI3753 Team\20 - Trojan Submittals\1 - Shop Dwg'
groutarea_list = os.listdir(GR_saved_path)

master_area_list = []
master_folder_list = []
master_file_list = []


for area in groutarea_list:
    if "advance" in area.lower():
        continue
    if "C-" not in area:
        continue
    print(area)
    folder_path = f"{GR_saved_path}\{area}"
    folder_list = os.listdir(folder_path)
    for file in folder_list:
        if not file.endswith(".pdf"):
            continue
        if "naco" in file.lower():
            file_path = f"{folder_path}\{file}"
            print(file_path)

            master_area_list.append(area)
            master_folder_list.append(file.replace(".pdf", ""))
            master_file_list.append(file_path)
        #     continue
        # if "reduce" in file.lower():
        #     continue
        # file_path = f"{folder_path}\{file}"
        # print(file_path)
        #
        # master_area_list.append(area)
        # master_folder_list.append(file.replace(".pdf", ""))
        # master_file_list.append(file_path)
        # continue


outputList1 = list(zip(master_area_list, master_folder_list, master_file_list))
dataframe = pd.DataFrame(outputList1)
pd.set_option('display.max_rows', 5000)
pd.set_option('display.max_columns', 5000)
pd.set_option('display.width', 10000)
print(dataframe)

dataframe.to_csv('csv\BI3753_Trojan Shop Drawing List_naco_full_path_230503.csv')