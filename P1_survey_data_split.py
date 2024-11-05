import xlwings as xw

source_path = r"C:\Users\921722\Box\BI3753 Team\30 - Geotech\03-Level Surveys\Pier 1 Level surveys\20230127\Ground Level Monitoring (GH 01-GH02) 27.01.2023_NACO.xlsx"
WB = xw.Book(source_path)
WS = WB.sheets['Plot Date Timeframe']

for i in range(1,156):
    string = WS.range(f'A{i}').value
    string_split = string.split(" ")
    for j in range(0, 3):
        WS.range(f'B{i}').offset(0, j).value = string_split[j]
