import xlwings as xw

source_path = r"C:\Users\921722\Box\PC4559 Eiffage Gibraltar coastal work WIP\02_GIR and Geotechnical Tasks\BH Logs Data Extraction\PSD\Vibrocore_PSD.xlsx"
WB = xw.Book(source_path)
WS = WB.sheets['vibrocore_PSD']

for i in range(4, 95):
    string = WS.range(f'B{i}').value
    #p1 = string.split("-")[1].split(" ")[0].strip()
    p2 = string.split("-")[1].split()[1].strip()
    print(string.split("-")[1])
    #WS.range(f'D{i}').value = p1
    WS.range(f'E{i}').value = p2

