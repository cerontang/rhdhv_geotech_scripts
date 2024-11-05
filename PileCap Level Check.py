import xlwings as xw
import numpy as np
import csv
import pandas as pd

with open(r'C:\Users\921722\OneDrive - Royal HaskoningDHV\20221107 AD\All Grouting Record Coordinates.csv', mode='r', encoding='utf-8-sig') as csvfile:
    data = [(str(n), float(x), float(y)) for n, x, y in csv.reader(csvfile, delimiter= ',')]
    csvfile.close()

data_coordinate_only = []
for element in data:
    coordinate_set = []
    for m in range(1,3):
        coordinate_set.append(element[m])
    data_coordinate_only.append(coordinate_set)

XL_path = r'C:\Users\921722\OneDrive - Royal HaskoningDHV\20221107 AD\pilecap_level.xlsx'
WB = xw.Book(XL_path)
WS_pcPoints = WB.sheets["Location R2"]
WS_pcLevels = WB.sheets["Levels"]

DATA = np.array(data_coordinate_only, dtype="float64")

n = 2
outputList = []

for i in range(2,3549):
    VBH_coordinate = np.array((float(WS_pcPoints.range(f"J{i}").value), float(WS_pcPoints.range(f"K{i}").value)))
    probe_dictionary = {"pilecap_level": [], "distanceBase": []}
    outputList_local = [WS_pcPoints.range(f"C{i}").value]
    distances = np.linalg.norm(DATA - VBH_coordinate, axis=1)
    min_distances_indices = np.argpartition(distances, n)[:n]
    min_distances = distances[min_distances_indices]
    print(distances)
    # for j in range(0,n):
    #     outputList_local_local = [data[min_distances_indices[j]][0]]
    #     outputList_local_local.append(min_distances[j])
    #     outputList_local.append(outputList_local_local)
    for j in range(0, n):
        outputList_local.append(data[min_distances_indices[j]][0])
        outputList_local.append(min_distances[j])
        probe_dictionary["pilecap_level"].append(data[min_distances_indices[j]][0])
        probe_dictionary["distanceBase"].append(min_distances[j])

    print(zip(list(probe_dictionary["pilecap_level"]), list(probe_dictionary["distanceBase"])))
    FINAL_zip = [WS_pcPoints.range(f"C{i}").value]
    for x in sorted(zip(list(probe_dictionary["distanceBase"]), list(probe_dictionary["pilecap_level"]))):
        for y in range(len(x)):
            FINAL_zip.append(x[y])

    print(FINAL_zip)
    outputList.append(FINAL_zip)

df = pd.DataFrame(outputList)
pd.set_option('display.max_rows', 5000)
pd.set_option('display.max_columns', 5000)
pd.set_option('display.width', 10000)
print(df)

df.to_csv(r'C:\Users\921722\OneDrive - Royal HaskoningDHV\20221107 AD\PileCapLevels_220911_2.csv')