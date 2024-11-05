from difflib import SequenceMatcher
import numpy as np
import csv
import pandas as pd
import xlwings as xw

XL_path = r'C:\Users\921722\Box\BI3753 Team\30 - Geotech\06-Grouting works\02 - Verification\04 OpenGround\230502_New VBH Info csvs\Stratum Detail Descriptions.csv'
WB = xw.Book(XL_path)
WS = WB.sheets["Stratum Detail Descriptions"]

for i in range (2, 110):
    description = WS.range(f"D{i}").value
    description = description.split("from")[0].strip()
    description = description.split(" at ")[0].strip()
    if not description.endswith("."):
        description += "."
    print(description)
    WS.range(f"D{i}").value = description


