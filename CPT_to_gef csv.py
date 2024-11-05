import xlwings as xw

file_path = r"C:\Users\921722\Royal HaskoningDHV\P-PC3110-Quays-land - Team\WIP\PC3110 Quays land WIP\Geotechnical\CPT investigation\GEF Conversion\20240704_Merged CPT data for GEF Conversion.xlsx"
WB = xw.Book(file_path)
WS_data = WB.sheets["Merged Data"]
WS_row = WB.sheets["Row Count"]
WS_csv = WB.sheets["csv"]

headers = [
    ["depth", "level", "Qc", "fs", "u", "Rf"],
    ["(m)", "(m)", "(MPa)", "(MPa)", "(MPa)", "(%)"]
]

columns = ["E", "B", "F", "G", "I", "AD"]

first_entry = True


for i in range(2, 35):
    ID = WS_row.range(f"A{i}").value
    row = int(WS_row.range(f"C{i}").value)
    start = int(WS_row.range(f"D{i}").value)
    end = int(WS_row.range(f"E{i}").value)
    print(row)

    if first_entry:
        WS_csv.range(f"A{row}").value = ID
        WS_csv.range(f"A{row + 1}").value = headers[0]
        WS_csv.range(f"A{row + 2}").value = headers[1]
        first_entry = False
        for j, col in enumerate(columns):
            WS_csv.range(f"A{row + 3}").offset(0, j).formula = f"='Merged Data'!{col}{start}:{col}{end}"
        continue

    WS_csv.range(f"A{row}").value = ID
    WS_csv.range(f"A{row + 2}").value = headers[0]
    WS_csv.range(f"A{row + 3}").value = headers[1]
    for j, col in enumerate(columns):
        WS_csv.range(f"A{row + 4}").offset(0, j).formula = f"='Merged Data'!{col}{start}:{col}{end}"

WB.app.calculate()