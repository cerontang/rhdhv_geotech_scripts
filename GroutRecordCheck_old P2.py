import os
import pandas as pd

GR_saved_path = r'C:\Users\921722\Box\BI3753 Team\30 - Geotech\06-Grouting works\01 - Grouting Records\Pier 2\Pier 2 Grouting Records Received 19-04-2022'
groutarea_list = os.listdir(GR_saved_path)

master_area_list = []
master_folder_list = []
master_file_list = []


for area in groutarea_list:
    if area == "Additional":
        continue
    folder_path = f"{GR_saved_path}\{area}"
    files = os.listdir(folder_path)
    for file in files:
        if file is None:
            break
        print(file)
        if not file.endswith(".pdf"):
            continue
        master_area = file.split("-")[0] + "-" + file.split("-")[1]
        gpID = file.split("-")[0] + "-" + file.split("-")[1] + "-" + file.split("-")[2]
        try:
            gpID = gpID.replace(".pdf", "")
        except:
            pass
        master_area_list.append(master_area)
        master_folder_list.append(gpID)
        master_file_list.append(file)


outputList1 = list(zip(master_area_list, master_folder_list, master_file_list))
dataframe = pd.DataFrame(outputList1)
pd.set_option('display.max_rows', 5000)
pd.set_option('display.max_columns', 5000)
pd.set_option('display.width', 10000)
print(dataframe)

dataframe.to_csv('csv\pier2part1_230322.csv')