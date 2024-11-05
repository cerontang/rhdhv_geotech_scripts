import os
import xlwings as xw
import pandas as pd
import matplotlib.pyplot as plt
import datetime

file_path = r"C:\Users\921722\Box\BI3753 Team\30 - Geotech\03-Level Surveys\Landside Surveys\Landside Survey Level Discrepancies for CMA clarifications.xlsx"
WB = xw.Book(file_path)
WS = WB.sheets["landside Storm"]

for i in range(2, 59):
    diff = abs(WS.range(f"M{i}").value)
    print(diff)
    if diff == 0:
        WS.range(f"L{i}").value = "0"
        continue
    if diff < 10:
        WS.range(f"L{i}").value = "< 10"
        continue
    if diff >= 10 and diff < 20:
        WS.range(f"L{i}").value = "10 - 20"
        continue
    if diff >= 20 and diff < 30:
        WS.range(f"L{i}").value = "20 - 30"
        continue
    if diff >= 30 and diff < 40:
        WS.range(f"L{i}").value = "30 - 40"
        continue
    if diff >= 40 and diff < 50:
        WS.range(f"L{i}").value = "40 - 50"
        continue
    if diff >= 50 and diff < 100:
        WS.range(f"L{i}").value = "50 - 100"
        continue
    if diff >= 100 and diff < 200:
        WS.range(f"L{i}").value = "100 - 200"
        continue
    if diff >= 200 and diff < 300:
        WS.range(f"L{i}").value = "200 - 300"
        continue
    if diff >= 300 and diff < 400:
        WS.range(f"L{i}").value = "300 - 400"
        continue
    if diff >= 400 and diff < 500:
        WS.range(f"L{i}").value = "400 - 500"
        continue
    if diff >= 500 and diff < 1000:
        WS.range(f"L{i}").value = "500 - 1000"
        continue
    if diff >= 1000:
        WS.range(f"L{i}").value = "> 1000"
        continue