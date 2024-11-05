import os
import xlwings as xw
import pandas as pd
import matplotlib.pyplot as plt
import datetime

file_path = r"C:\Users\921722\Box\BI3753 Team\30 - Geotech\03-Level Surveys\Landside Surveys\Landside area Settlement Monitoring_R1_CT.xlsx"
WB = xw.Book(file_path)
WS = WB.sheets["Landside Pavement Master R1"]

for i in range (3, 114):
    #date_list = []
    survey_list = []
    for j in range (0, 34):
        if WS.range(f"AT{i}").offset(0, j).value == '':
            continue
        survey_list.append(round(WS.range(f"AT{i}").offset(0, j).value, 3))
        #date_list.append(WS.range(f"AT2").offset(0, j).value)
    #print(date_list)
    Settlement = 'N'
    Heave = 'N'
    for survey in survey_list:
        if survey == 0:
            continue
        if survey < 0:
            Settlement = 'Y'
        if survey > 0:
            Heave = 'Y'
    if Settlement == 'N' and Heave == 'N':
        WS.range(f"CD{i}").value = "No displacement"
        continue
    if Settlement == 'Y' and Heave == 'Y':
        WS.range(f"CD{i}").value = "Settlement + Heave"
        continue
    if Settlement == 'Y' and Heave == 'N':
        WS.range(f"CD{i}").value = "Settlement"
        continue
    if Settlement == 'N' and Heave == 'Y':
        WS.range(f"CD{i}").value = "Heave"
        continue
    #print(i - 2, survey_list)