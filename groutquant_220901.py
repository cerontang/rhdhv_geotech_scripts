import xlwings as xw
import pandas as pd
import matplotlib.pyplot as plt
import datetime

file_path = r"C:\Users\921722\OneDrive - Royal HaskoningDHV\20230602 AD\folder\ERI_Descriptions_02-06-2023.xlsx"
WB = xw.Book(file_path)
WS = WB.sheets["ERI_MAIN"]

for i in range (2, 500):
    description = WS.range(f"E{i}").value
    description_split = description.split("%")
    if len(description_split) < 2:
        continue
    description_quant = description_split[0]
    quant = description_quant.split()[-1].strip()
    WS.range(f"G{i}").value = quant

