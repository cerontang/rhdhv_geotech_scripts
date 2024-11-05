import xlwings as xw
from statistics import mean

file_path = r"C:\Users\921722\Box\BI3753 Team\30 - Geotech\03-Level Surveys\Landside Surveys\Landside area Settlement Monitoring_R1_CT.xlsx"
WB = xw.Book(file_path)
WS = WB.sheets["Landside Storm Master R1"]

for i in range (3, 330):
    local_IL_list = []
    if WS.range(f"AC{i}").value is None:
        for j in range (1, 10):
            if WS.range(f"D{i+j}").value == "MH IL":
                continue
            if WS.range(f"AC{i+j}").value is None:
                break
            local_IL_list.append(WS.range(f"AC{i+j}").value)
    try:
        average = mean(local_IL_list)
        WS.range(f"AB{i}").value = average
    except:
        pass