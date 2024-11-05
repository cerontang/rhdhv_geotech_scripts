from difflib import SequenceMatcher
import numpy as np
import csv
import pandas as pd
import xlwings as xw

XL_path = r'C:\Users\921722\Box\BI3753 Team\30 - Geotech\06-Grouting works\01 - Grouting Records\Check 230412\Grouting Record Check_230412.xlsx'
WB = xw.Book(XL_path)
WS = WB.sheets["pier2_230412"]

for i in range (2, 5424):
    description = WS.range(f"D{i}").value
    if "permeation" in description.lower():
        WS.range(f"F{i}").value = "Permeation"
        continue
    if "block" in description.lower():
        WS.range(f"F{i}").value = "Block"
        continue
    if "drilling" in description.lower():
        WS.range(f"F{i}").value = "Drilling"
        continue

    # dash_count = description.count("-")
    # if dash_count == 4:
    #     WS.range(f"F{i}").value = "Permeation"
    #     continue
    # if dash_count == 3:
    #     WS.range(f"F{i}").value = "Block"
    #     continue
    # if dash_count == 2:
    #     WS.range(f"F{i}").value = "Drilling"
    #     continue



# for i in range (2, 4606):
#     description = WS.range(f"E{i}").value
#     if "-N-" in description:
#         description = description.split("-")[0] + "-" + description.split("-")[1] + "-" + description.split("-")[2]
#         WS.range(f"D{i}").value = description
#     else:
#         description = description.split("-")[0] + "-" + description.split("-")[1]
#         WS.range(f"D{i}").value = description
#     print(description)




