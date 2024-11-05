import xlwings as xw
import pandas as pd
import numpy as np

file_path = r"C:\Users\921722\OneDrive - Royal HaskoningDHV\20220908 AD\Priority 1 task\Pier 1 Grout Completion Status - Pivot Table.xlsx"
WB = xw.Book(file_path)
WS = WB.sheets["Summary"]
#WS2 = WB.sheets["Output Table 220908"]

count = 0
section = "LS12-S1301"

for i in range(2, 2814):
    if WS.range(f"F{i}").value != "INCOMPLETE":
        continue
    print(WS.range(f"E{i}").value, WS.range(f"B{i}").value, WS.range(f"F{i}").value)
    if WS.range(f"B{i}").value == section and WS.range(f"F{i}").value == "INCOMPLETE":
        probeName = WS.range(f"E{i}").value
        if "SEW" in probeName:
            count += 1

print(count)

# for i in range(2, 2814):
#     probeName = WS.range(f"E{i}").value
#     print(WS.range(f"E{i}").value, WS.range(f"B{i}").value, WS.range(f"F{i}").value)
#     if "PR01" in probeName and WS.range(f"F{i}").value == "INCOMPLETE":
#         count += 1
#
# print(count)