import xlwings as xw
import pandas as pd
import matplotlib.pyplot as plt
import datetime
import numpy as np

master_VBH_list = []
master_depth_list = []

file_path = r"C:\Users\921722\OneDrive - Royal HaskoningDHV\20220902AD\VBH_Grout Information_20220816_check_2.xlsx"
WB = xw.Book(file_path)
WS = WB.sheets["Sheet1"]

for i in range(1, 70):
    VBH_name = str(WS.range(f"A{i}").value)
    max_depth = WS.range(f"C{i}").value + 0.5
    depth_list = np.arange(0, max_depth, 0.5).tolist()
    print(depth_list)
    for j in range(len(depth_list)):
        master_VBH_list.append(VBH_name)
        master_depth_list.append(depth_list[j])
        print(VBH_name, depth_list[j])


print(master_VBH_list, master_depth_list)
output_list = list(zip(master_VBH_list, master_depth_list))
df = pd.DataFrame(output_list)
df.to_csv(r"csv\grout FUGRO 2 0.5 220902.csv")
print(df)