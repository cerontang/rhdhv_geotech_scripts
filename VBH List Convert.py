import xlwings as xw

XL_path = r'C:\Users\921722\Box\BI3753 Team\30 - Geotech\06-Grouting works\02 - Verification\07 Summaries\VBH Priority List 3-Pier 2 and 1_221004.xlsx'
WB = xw.Book(XL_path)
WS = WB.sheets["Sheet1"]

for i in range (4, 34):
    try:
        description = WS.range(f"I{i}").value
        description_remarks = description.split(". Indicative position")[0].strip() + "."
        description_coordinates = description.split(". Indicative position")[1].strip()
        #WS.range(f"F{i}").value = description
        description_coordinates = description_coordinates.split(" ")
        WS.range(f"F{i}").value = description_coordinates[1]
        WS.range(f"G{i}").value = description_coordinates[3]
        #WS.range(f"I{i}").value = description_remarks
        print(description_coordinates, description_remarks)
    except:
        continue

