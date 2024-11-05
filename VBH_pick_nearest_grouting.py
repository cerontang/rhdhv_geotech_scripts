import xlwings as xw
import numpy as np
import csv
import pandas as pd
import datetime

with open(r'C:\Users\921722\OneDrive - Royal HaskoningDHV\20220817 AD\Grouting Records\All Grouting Record Coordinates.csv', mode='r', encoding='utf-8-sig') as csvfile:
    data = [(str(n), float(x), float(y)) for n, x, y in csv.reader(csvfile, delimiter= ',')]
    csvfile.close()

data_coordinate_only = []
for element in data:
    coordinate_set = []
    for m in range(1,3):
        coordinate_set.append(element[m])
    data_coordinate_only.append(coordinate_set)

XL_path = r'C:\Users\921722\OneDrive - Royal HaskoningDHV\20220817 AD\Grouting Records\ALL VBH Coordinates.xlsx'
WB = xw.Book(XL_path)
WS = WB.sheets["Sheet1"]

DATA = np.array(data_coordinate_only, dtype="float64")
# n closest grout point
n = 4
outputList = []

for i in range(2,356):
    outputList_local = [WS.range(f"D{i}").value]
    VBH_coordinate = np.array((WS.range(f"F{i}").value, WS.range(f"G{i}").value))
    distances = np.linalg.norm(DATA - VBH_coordinate, axis=1)
    min_distances_indices = np.argpartition(distances, n)[:n]
    min_distances = distances[min_distances_indices]
    # for j in range(0,n):
    #     outputList_local_local = [data[min_distances_indices[j]][0]]
    #     outputList_local_local.append(min_distances[j])
    #     outputList_local.append(outputList_local_local)
    for j in range(0, n):
        outputList_local.append(data[min_distances_indices[j]][0])
        outputList_local.append(min_distances[j])
    print(outputList_local)
    outputList.append(outputList_local)

df = pd.DataFrame(outputList)
pd.set_option('display.max_rows', 5000)
pd.set_option('display.max_columns', 5000)
pd.set_option('display.width', 10000)
print(df)
df.to_csv(f'csv\VBH_groutpoints_220831.csv')

