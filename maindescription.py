import os
import xlwings as xw
import pandas as pd
import matplotlib.pyplot as plt
import datetime

file_path = r"C:\Users\921722\Royal HaskoningDHV\P-PC6298-Hunterston - Team\WIP\4. Geotechnical\02. OpenGround\1. Inputs\3. Borehole Extraction\FGE.xlsx"
WB = xw.Book(file_path)
WS = WB.sheets["data"]

# for i in range(2, 677):
#     description = WS.range(f"E{i}").value
#     split_description = description.split()
#     for word in split_description:
#         if word.isupper():
#             mainWord = word
#             break
#     mainWord = mainWord.replace('.', "")
#     mainWord = mainWord.replace(',', "")
#     mainWord = mainWord.strip()
#     WS.range(f"F{i}").value = mainWord

for i in range(2, 1000):
    description_capital_list = []
    description = WS.range(f"D{i}").value
    split_description = description.split()
    for word in split_description:
        if not word.isupper():
            continue
        word = word.replace('.', "")
        word = word.replace(',', "")
        word = word.replace(')', "")
        word = word.replace('())', "")
        description_capital_list.append(word)
        continue
    if len(description_capital_list) > 1:
        #mainWord = description_capital_list[len(description_capital_list)-1]
        mainWord = description_capital_list[-1]
        WS.range(f"S{i}").value = "More than one"
    else:
        if len(description_capital_list) == 0:
            continue
        mainWord = description_capital_list[-1]

    WS.range(f"R{i}").value = mainWord

