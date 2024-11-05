import xlwings as xw
import pandas as pd
import numpy as np

file_path = r"C:\Users\921722\Box\BI3753 Team\30 - Geotech\06-Grouting works\02 - Verification\04 OpenGround\230301_New VBH Info csvs\Stratum Detail Descriptions.csv"
WB = xw.Book(file_path)
WS = WB.sheets["Stratum Detail Descriptions"]

#remove newlines and trim text string
# for i in range(2, 34):
#     raw_description = WS.range(f"A{i}").value
#     description = raw_description.replace("\n", " ").strip()
#     WS.range(f"B{i}").value = description

for i in range (2, 30):
    description = WS.range(f"D{i}").value
    description_fixed = description.split("from")[0]
    description_fixed = description_fixed.split(" at ")[0]
    if not description_fixed.endswith("."):
        description_fixed = description_fixed.strip() + "."
    WS.range(f"D{i}").value = description_fixed