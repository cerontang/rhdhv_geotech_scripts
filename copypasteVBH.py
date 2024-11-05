import shutil
import xlwings as xw
import os

file_path = r"C:\Users\921722\OneDrive - Royal HaskoningDHV\20221005 AD\VBH\folder\AHAM\VBH_AHAM_Descriptions_220923.xlsx"
WB = xw.Book(file_path)
WS = WB.sheets["Sheet5"]

link_list = []
VBH_name_list = []
for i in range(1, 21):
    link_list.append(WS.range(f"B{i}").value)
    VBH_name_list.append(WS.range(f"A{i}").value)

combined_list = list(zip(VBH_name_list, link_list))

for item in combined_list:
    original_path = item[1]
    new_path = f"C:/Users/921722/OneDrive - Royal HaskoningDHV/20220902AD/VBH to be extracted/{item[0]}.pdf"
    shutil.copyfile(original_path, new_path)
