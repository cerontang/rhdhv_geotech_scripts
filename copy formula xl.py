import xlwings as xw

XL_path = r'C:\Users\921722\OneDrive - Royal HaskoningDHV\20220825 AD\P3 All Fugro VBH\folder\P3 Fugro VBH List 220825.xlsx'
WB = xw.Book(XL_path)
WS = WB.sheets["Sheet1"]

VBH_name_list = []

for a in range (2, 177):
    VBH_name = WS.range(f"C{a}").value
    VBH_name_list.append(VBH_name)
    if a == 2:
        continue
    if not VBH_name == VBH_name_list[a-3]:
        WS.range(f"E{a}").value = "i"

