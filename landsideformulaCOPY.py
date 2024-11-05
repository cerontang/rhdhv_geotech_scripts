import os
import xlwings as xw
import pandas as pd

file_path = r"C:\Users\921722\Box\BI3753 Team\30 - Geotech\03-Level Surveys\Landside Surveys\Landside area Settlement Monitoring_R1_CT.xlsx"
WB = xw.Book(file_path)
WS = WB.sheets["Landside Sewer Master"]

formulaOG = "=IF(ISBLANK(INDEX('Raw Data Merged June 2022'!$A$6:$CG$286,2+MATCH('Landside Sewer Master'!$D$1&'Landside Sewer Master'!C6,'Raw Data Merged June 2022'!$B$8:$B$286&'Raw Data Merged June 2022'!$D$8:$D$286,0), 13)), \"\", INDEX('Raw Data Merged June 2022'!$A$6:$CG$286,2+MATCH('Landside Sewer Master'!$D$1&'Landside Sewer Master'!C6,'Raw Data Merged June 2022'!$B$8:$B$286&'Raw Data Merged June 2022'!$D$8:$D$286,0), 14))"
count = 0

for i in range (36):
    formula = f"=IF(ISBLANK(INDEX('Raw Data Merged June 2022'!$A$6:$CG$286,2+MATCH('Landside Sewer Master'!$D$1&'Landside Sewer Master'!C6,'Raw Data Merged June 2022'!$B$8:$B$286&'Raw Data Merged June 2022'!$D$8:$D$286,0), {str(14+count)})), \"\", INDEX('Raw Data Merged June 2022'!$A$6:$CG$286,2+MATCH('Landside Sewer Master'!$D$1&'Landside Sewer Master'!C6,'Raw Data Merged June 2022'!$B$8:$B$286&'Raw Data Merged June 2022'!$D$8:$D$286,0)," + str(f" {14+count}))")
    WS.range('K6').offset(0, i).value = formula
    print(formula)
    count += 2

