import os
import pandas as pd

GR_saved_path = r'C:\Users\921722\Royal HaskoningDHV\Project-PC5465-BELFAST-D3 - Team\WIP\01 Geotechnical'
groutarea_list = os.listdir(GR_saved_path)

master_area_list = []
master_folder_list = []
master_file_list = []


for area in groutarea_list:
    if area == "Additional":
        continue
    folder_path = f"{GR_saved_path}\{area}"
    if folder_path.endswith(".pdf"):
        continue
    folder_list = os.listdir(folder_path)
    for folder in folder_list:
        file_path = f"{folder_path}\{folder}"
        try:
            file_list = os.listdir(file_path)
        except:
            master_area_list.append(area)
            master_folder_list.append(folder)
            master_file_list.append("")
            continue

        print(file_list)
        for file in file_list:
            master_area_list.append(area)
            master_folder_list.append(folder)
            master_file_list.append(file)

outputList1 = list(zip(master_area_list, master_folder_list, master_file_list))
dataframe = pd.DataFrame(outputList1)
pd.set_option('display.max_rows', 5000)
pd.set_option('display.max_columns', 5000)
pd.set_option('display.width', 10000)
print(dataframe)

dataframe.to_csv('csv\PC5465_Belfast D3_Analysis File Register.csv')