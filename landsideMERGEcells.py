import os
import xlwings as xw
import pandas as pd
import matplotlib.pyplot as plt

file_path = r"C:\Users\921722\Box\BI3753 Team\30 - Geotech\03-Level Surveys\Landside Surveys\Landside area Settlement Monitoring_R1_CT.xlsx"
WB = xw.Book(file_path)
WS = WB.sheets["Landside Pavement Master R1"]

# for i in range (3, 60):
#     trend_list = []
#     for j in range (0, 3):
#         #WS.range(f"O{i}").offset(0, j).merge()
#         trend_list.append(WS.range(f"O{i}").offset(0, j).value)
#     print(trend_list)

# for i in range (3, 60):
#     trend_list = []
#     for j in range (0, 3):
#         #WS.range(f"O{i}").offset(0, j).merge()
#         trend_list.append(WS.range(f"O{i}").offset(0, j).value)
#     if trend_list[1] is None or trend_list[2] is None:
#         continue
#     if trend_list[2] < trend_list[1]:
#         WS.range(f"U{i}").value = "Settlement"
#         continue
#     if trend_list[2] == trend_list[1]:
#         WS.range(f"U{i}").value = "No displacement"
#         continue
#     if trend_list[2] > trend_list[1]:
#         WS.range(f"U{i}").value = "Heave"
#         continue

# for i in range (11, 58):
#     trend_list = []
#     trend_list.append(WS.range(f"BZ{i}").value)
#     trend_list.append(WS.range(f"CE{i}").value)
#     print(trend_list)
#     if trend_list[0] == '' or trend_list[1] == '':
#         continue
#     if trend_list[1] < trend_list[0]:
#         WS.range(f"CH{i}").value = "Settlement"
#         continue
#     if trend_list[1] == trend_list[0]:
#         WS.range(f"CH{i}").value = "No displacement"
#         continue
#     if trend_list[1] > trend_list[0]:
#         WS.range(f"CH{i}").value = "Heave"
#         continue


