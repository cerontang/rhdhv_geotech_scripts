import pandas as pd

file = open(r"C:\Users\921722\OneDrive - Royal HaskoningDHV\20230112 AD\QGIS\LayerList_MGD.txt", "r")
input = file.readlines()
output = []

for item in input:
    item = str(item.replace("\n", "")).split("/")[-1]
    output.append(item)
    print(item)

dataframe = pd.DataFrame(output)
pd.set_option('display.max_rows', 5000)
pd.set_option('display.max_columns', 5000)
pd.set_option('display.width', 10000)
print(dataframe)

dataframe.to_csv(f'csv\QGISLayer.csv')