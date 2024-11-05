import xlwings as xw

FGV_path = r'C:\Users\921722\Box\BI3753 Team\30 - Geotech\05-Ground Investigation\03 Factual Reports\APM HH 2022\05 OpenGround\Field Geological Descriptions.csv'
WB_FGV = xw.Book(FGV_path)
WS_FGV = WB_FGV.sheets["Field Geological Descriptions"]

X_path = r'C:\Users\921722\OneDrive - Royal HaskoningDHV\20220729 AD\Extract Logs\Lab Test data separated\APM HH 2022 GI LAB TEST DATA CSVs\Liquid and Plastic Limit Tests.csv'
WB_X = xw.Book(X_path)
WS_X = WB_X.sheets["Liquid and Plastic Limit Tests"]

for i in range(2, 15):
    BH_name = WS_X.range(f"A{i}").value
    d_t = WS_X.range(f"H{i}").value
    print(d_t)
    for j in range(2, 339):
        if BH_name != WS_FGV.range(f"A{j}").value:
            continue
        if d_t >= WS_FGV.range(f"C{j}").value and d_t < WS_FGV.range(f"D{j}").value:
            WS_X.range(f"B{i}").value = WS_FGV.range(f"C{j}").value
            WS_X.range(f"M{i}").value = WS_FGV.range(f"I{j}").value
            break