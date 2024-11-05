import xlwings as xw
import pandas as pd


asBuilt_file_path = f'C:/Users/921722/Box/BI3753 Team/30 - Geotech/03-Level Surveys/Incoming/20220513-Storm Manholes/AS built storm water MAY 2022 Invert level and cover level Airside All piers - NACO.xlsx'
asBuilt_WB = xw.Book(asBuilt_file_path)
asBuilt_WS = asBuilt_WB.sheets["Design IL Received"]



for i in range(3, 472):
    if asBuilt_WS.range(f'B{i}').value is None:
        asBuilt_WS.range(f'B{i}').value = asBuilt_WS.range(f'B{i-1}').value


