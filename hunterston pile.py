import xlwings as xw

file_path = r"C:\Users\921722\Royal HaskoningDHV\P-PC6298-Hunterston - Team\WIP\4. Geotechnical\07. Pile Design\script\Pile Bearing Capacity Calcs_for script.xlsm"
WB = xw.Book(file_path)
WS = WB.sheets["Pile Toe Levels_150kPa"]

spacing_list = [(3.5, 5), (7, 5)]
pile_list = [0.9, 1.05, 1.2, 1.35, 1.5, 1.65, 1.83]

for x, y in spacing_list:
    WS.range(f"B3").value = x
    WS.range(f"B4").value = y
    for d in pile_list:
        WS.range(f"B11").value = d
        for i in range(-20, -80, -1):
            WS.range(f"B16").value = i
            # if i == -44 and WS.range(f"B38").value / WS.range(f"B9").value < 1.0:
            #     local_list = [d, x, y, i, WS.range(f"B9").value, WS.range(f"B38").value]
            #     print(local_list)
            #     break
            if WS.range(f"B38").value / WS.range(f"B9").value >= 1.0:
                local_list = [d, x, y, i, WS.range(f"B9").value, WS.range(f"B38").value]
                print(local_list)
                break