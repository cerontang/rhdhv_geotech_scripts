import xlwings as xw

file_path = r"C:\Users\921722\Box\BI3753 Team\99 - Temp\_SSCT\20230503 Trojan Shop Drawing Check\Version 2\BI3753_Trojan Shop Drawing List_230503.xlsx"
WB = xw.Book(file_path)
WS = WB.sheets["SDS image list"]

for i in range (2, 631):
    image_name = str(WS.range(f"E{i}").value)
    WS.range(f"D{i}").value = image_name.split("_Page")[0]

