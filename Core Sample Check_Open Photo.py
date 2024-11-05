import os
import xlwings as xw

XL_path = r'C:\Users\921722\Box\BI3753 Team\99 - Temp\_SSCT\20230307 - Corehole UCS Samples\MTB_Corehole UCS Samples_230307.xlsx'
WB = xw.Book(XL_path)
WS = WB.sheets["master"]
WS2 = WB.sheets["photos"]

rowNumber = "2"
VBH_Name = str(WS.range(f"F{rowNumber}").value)

path = f"C:/Users/921722/Box/BI3753 Team/99 - Temp/_SSCT/20230307 - Corehole UCS Samples/photos"
if not os.path.exists(f"{path}/{VBH_Name}"):
    pass
ph_path = f"{path}/{VBH_Name}"
ph_list = os.listdir(ph_path)
for row, file in list(enumerate(ph_list)):
    open_path = f"{ph_path}/{file}"
    os.startfile(open_path)


