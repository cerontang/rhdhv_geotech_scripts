import os
import xlwings as xw
import pandas as pd
import matplotlib.pyplot as plt
import datetime

file_path = r"C:\Users\921722\OneDrive - Royal HaskoningDHV\20220818 AD\APMHH2022_GI_Description_AHAM_220817.csv"
WB = xw.Book(file_path)
WS = WB.sheets["Field Geological Descriptions"]

for i in range(1, 29):
    description = WS.range(f"E{i}").value
    split_description = description.split()
    indexed_split_description_list = list(enumerate(split_description))
    print(indexed_split_description_list)
    for word in indexed_split_description_list:
        grout_exist = "N"
        #print(word)
        if word[1] != "grout":
            continue
        if word[1] == "grout":
            grout_exist = "Y"
            groutValue = indexed_split_description_list[word[0]-2][1].replace("%", "")
            WS.range(f"G{i}").value = groutValue
            break



