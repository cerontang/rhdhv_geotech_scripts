import os
import xlwings as xw

XL_path = r'C:\Users\921722\Box\BI3753 Team\99 - Temp\_SSCT\20230307 - Corehole UCS Samples\MTB_Corehole UCS Samples_230307.xlsx'
WB = xw.Book(XL_path)
WS = WB.sheets["master"]
WS2 = WB.sheets["photos"]

rowNumber = 2
VBH_Name = str(WS.range(f"F{rowNumber}").value)

path = f"C:/Users/921722/Box/BI3753 Team/99 - Temp/_SSCT/20230307 - Corehole UCS Samples/photos"
if not os.path.exists(f"{path}/{VBH_Name}"):
    print("file not found")
else:
    ph_path = f"{path}/{VBH_Name}"
    ph_list = os.listdir(ph_path)
    row = 1  # starting row for photos
    for file in ph_list:
        open_path = f"{ph_path}/{file}"
        while not WS2.range(f"A{row+1}").value is None:
            row += 1
        WS2.pictures.add(open_path, name=f"Photo_{row}", left=WS.range("A1").left, top=WS.range(f"A{row}").top, width=200, height=200)
        row += 1