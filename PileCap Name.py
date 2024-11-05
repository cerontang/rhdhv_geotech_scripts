import xlwings as xw

XL_path = r'C:\Users\921722\OneDrive - Royal HaskoningDHV\20221107 AD\PileCapLevels_220911_algo.xlsx'
WB = xw.Book(XL_path)
WS = WB.sheets["Final"]

for i in range (2, 1691):

    #FUNCTION 1

    # try:
    #     raw_name = WS.range(f"A{i}").value
    #     raw_name_split = raw_name.strip().split("&")
    #     raw_name_1 = raw_name_split[0]
    #     raw_name_2 = raw_name_split[1]
    #     print(raw_name_1, raw_name_2)
    #     WS.range(f"B{i}").value = raw_name_1
    #     WS.range(f"C{i}").value = raw_name_2
    # except:
    #     WS.range(f"B{i}").value = raw_name.strip()

    #FUNCTION 2

    # loc_list = [str(WS.range(f"D{i}").value), str(WS.range(f"E{i}").value)]
    # if loc_list[0] == "Central" and loc_list[1] == "Central":
    #     WS.range(f"F{i}").value = "MTB Central"
    #     continue
    # if (loc_list[0] == "Edge" and loc_list[1] == "Edge") or (loc_list[0] == "Central" and loc_list[1] == "Edge") or (loc_list[0] == "Edge" and loc_list[1] == "Central"):
    #     WS.range(f"F{i}").value = "MTB Edge"
    #     continue
    # if (loc_list[0] == "Other, Landside" and loc_list[1] == "Other, Landside") or (loc_list[0] == "Central" and loc_list[1] == "Other, Landside"):
    #     WS.range(f"F{i}").value = "Landside"
    #     continue
    # if loc_list[1] == "#N/A":
    #     WS.range(f"F{i}").value = loc_list[1]
    #     continue
    # WS.range(f"F{i}").value = "Landside Edge"



