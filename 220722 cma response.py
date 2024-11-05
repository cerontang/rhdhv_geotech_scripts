import xlwings as xw

file_path = r"C:\Users\921722\OneDrive - Royal HaskoningDHV\20220729 AD\APMHH2022_GI_Description_ACES.xlsx"
WB = xw.Book(file_path)
WS = WB.sheets["Sheet3"]

for i in range (2, 43):
    description = WS.range(f"D{i}").value
    description_list = description.split("*")
    if len(description_list) > 1:
        WS.range(f"E{i}").value = 'Y'
        continue