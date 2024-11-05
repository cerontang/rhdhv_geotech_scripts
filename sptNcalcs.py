import xlwings as xw

filePath = r'C:\Users\921722\Box\PC3348 Greenway Quay\PC3348 Greenway Quay WIP\Geotechnical\borehole calcs.xlsx'

WB = xw.Book(filePath)
WS = WB.sheets[f'Sheet1']

for j in range(1, 11):
    DPname = f"DP{j}"
    firstorNot = "Y"
    for i in range(4, 457):
        if WS.range(f'A{i}').value == DPname:
            if firstorNot == "Y":
                WS.range(f'D{i}').value = WS.range(f'C{i}').value + WS.range(f'C{i+1}').value + WS.range(f'C{i+2}').value
                firstorNot = "N"
                continue
            if WS.range(f'A{i}').value != WS.range(f'A{i+1}').value:
                WS.range(f'D{i}').value = WS.range(f'C{i}').value
                continue
            if WS.range(f'D{i-1}').value or WS.range(f'D{i-2}').value is not None:
                continue
            if WS.range(f'D{i-1}').value is None and WS.range(f'D{i-2}').value is None:
                WS.range(f'D{i}').value = WS.range(f'C{i}').value + WS.range(f'C{i+1}').value + WS.range(f'C{i+2}').value
                continue



