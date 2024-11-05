import xlwings as xw
import os

file_path = r"C:\Users\921722\Box\BI3753 Team\30 - Geotech\03-Level Surveys\Landside Surveys\Landside area Settlement Monitoring_R3_230206.xlsx"
WB = xw.Book(file_path)
WS = WB.sheets["Landside Storm Master R2"]


for i in range (4, 340):
    content = WS.range(f'V{i}').value
    if content is None:
        continue
    formula = f"=SERIES('Landside Storm Master R2'!$C${i},'Landside Storm Master R2'!$V$2:$Y$2,'Landside Storm Master R2'!$V${i}:$Y${i},1)"
    print(formula)