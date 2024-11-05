import xlwings as xw
import os

file_path = r"C:\Users\921722\Box\BI3753 Team\30 - Geotech\03-Level Surveys\Landside Surveys\Landside BH List vs Survey Assessment_230208.xlsx"
WB = xw.Book(file_path)
WS = WB.sheets["Landside BH List_R1_230130"]

for i in range(2,58):
    locationID = WS.range(f'E{i}').value
    MH_name = locationID.split("-")[1]
    WS.range(f'D{i}').value = MH_name