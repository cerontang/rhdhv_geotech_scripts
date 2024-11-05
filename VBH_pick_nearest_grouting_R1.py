import xlwings as xw
import numpy as np
import csv
import pandas as pd
from dataclasses import dataclass

with open(r'C:\Users\921722\Box\BI3753 Team\30 - Geotech\03-Level Surveys\Pier 1 Level surveys\20230127\script\All Grouting Record Coordinates.csv', mode='r', encoding='utf-8-sig') as csvfile:
    data = [(str(n), float(x), float(y)) for n, x, y in csv.reader(csvfile, delimiter= ',')]
    csvfile.close()

data_coordinate_only = []
for element in data:
    coordinate_set = []
    for m in range(1,3):
        coordinate_set.append(element[m])
    data_coordinate_only.append(coordinate_set)

XL_path = r'C:\Users\921722\Box\BI3753 Team\30 - Geotech\03-Level Surveys\Pier 1 Level surveys\20230127\Ground Level Monitoring (GH 01-GH02) 27.01.2023_NACO_macro.xlsx'
WB = xw.Book(XL_path)
WS = WB.sheets["Survey Points"]

DATA = np.array(data_coordinate_only, dtype="float64")
# n closest grout point
n = 6
outputList = []

for i in range(2,168):
    if WS.range(f"B{i}").value is None:
        continue
    probe_dictionary = {"probeID": [], "distanceBase": []}
    outputList_local = [WS.range(f"A{i}").value]
    VBH_coordinate = np.array((WS.range(f"B{i}").value, WS.range(f"C{i}").value))
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
        probe_dictionary["probeID"].append(data[min_distances_indices[j]][0])
        probe_dictionary["distanceBase"].append(min_distances[j])

    print(zip(list(probe_dictionary["probeID"]), list(probe_dictionary["distanceBase"])))
    FINAL_zip = [WS.range(f"A{i}").value]
    for x in sorted(zip(list(probe_dictionary["distanceBase"]), list(probe_dictionary["probeID"]))):
        for y in range(len(x)):
            FINAL_zip.append(x[y])

    print(FINAL_zip)
    outputList.append(FINAL_zip)

df = pd.DataFrame(outputList)
pd.set_option('display.max_rows', 5000)
pd.set_option('display.max_columns', 5000)
pd.set_option('display.width', 10000)
print(df)
df.to_csv(f'csv\survey.csv')


