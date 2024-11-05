import xlwings as xw


FGV_path = r'C:\Users\921722\Box\BI3753 Team\30 - Geotech\05-Ground Investigation\03 Factual Reports\APM HH 2022\05 OpenGround\Coring Information.csv'
WB_FGV = xw.Book(FGV_path)
WS_FGV = WB_FGV.sheets["Coring Information"]

FS_path = r'C:\Users\921722\Box\BI3753 Team\30 - Geotech\05-Ground Investigation\03 Factual Reports\APM HH 2022\05 OpenGround\Rock Uniaxial Compressive Strength and Deformability Tests.csv'
WB_FS = xw.Book(FS_path)
WS_FS = WB_FS.sheets["Rock Uniaxial Compressive Stren"]

for i in range(2, 45):
    BH_name = WS_FS.range(f"A{i}").value
    d_specimen = WS_FS.range(f"G{i}").value
    #print(d_specimen)
    for j in range(2, 195):
        if BH_name != WS_FGV.range(f"A{j}").value:
            continue
        if d_specimen >= WS_FGV.range(f"B{j}").value and d_specimen <= WS_FGV.range(f"C{j}").value:
            print(WS_FS.range(f"G{i}").value, WS_FGV.range(f"B{j}").value, WS_FS.range(f"G{i}").value >= WS_FGV.range(f"B{j}").value)
            WS_FS.range(f"B{i}").value = WS_FGV.range(f"B{j}").value
            break
