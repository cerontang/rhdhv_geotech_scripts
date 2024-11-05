import os
import xlwings as xw
import pandas as pd

rhdhv_file_path = f'C:/Users/921722/Box/BI3753 Team/30 - Geotech/03-Level Surveys/Sewer and Storm surveys/Stormwater Manhole monitoring-MASTER-R5.xlsx'
# asBuilt_file_path = f'C:/Users/921722/Box/BI3753 Team/30 - Geotech/03-Level Surveys/Incoming/20220513-Storm Manholes/AS built storm water MAY 2022 Invert level and cover level Airside All piers - NACO.xlsx'

rhdhv_WB = xw.Book(rhdhv_file_path)
rhdhv_WS = rhdhv_WB.sheets["Data-Invert"]
# asBuilt_WB = xw.Book(asBuilt_file_path)
# asBuilt_WS = asBuilt_WB.sheets["Sheet1"]

designIL_list = []
manholeID_list = []

for i in range(3, 563):
    if rhdhv_WS.range(f'D{i+1}').value == "CL":
        designIL_list.append("duplicate")
        manholeID_list.append(rhdhv_WS.range(f'D{i}').value)
        continue
    if rhdhv_WS.range(f'D{i}').value == "CL":
        designIL_list.append(rhdhv_WS.range(f'E{i}').value)
        manholeID_list.append(rhdhv_WS.range(f'D{i-1}').value)
        continue
    designIL_list.append(rhdhv_WS.range(f'E{i}').value)
    manholeID_list.append(rhdhv_WS.range(f'D{i}').value)

outputList = list(zip(manholeID_list, designIL_list))

df = pd.DataFrame(outputList)
pd.set_option('display.max_rows', 99999)
pd.set_option('display.max_columns', 99999)
pd.set_option('display.width', 99999)
print(df)
df.to_csv(r'csv\rhdhvIL.csv')




