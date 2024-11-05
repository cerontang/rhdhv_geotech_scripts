import xlwings as xw

source_path = r"C:\Users\921722\OneDrive - Royal HaskoningDHV\20230602 AD\folder\ERI_Descriptions_02-06-2023.xlsx"
WB = xw.Book(source_path)
WS = WB.sheets['ERI_MAIN']
WS2 = WB.sheets['ERI 0.5']
#index_list = []

# for i in range(2, 500):
#     description_capital_list = []
#     description = WS.range(f"E{i}").value
#     split_description = description.split()
#     for word in split_description:
#         if not word.isupper():
#             continue
#         word = word.replace('.', "")
#         word = word.replace(',', "")
#         word = word.replace(')', "")
#         word = word.replace('())', "")
#         description_capital_list.append(word)
#         continue
#     if len(description_capital_list) > 1:
#         mainWord = description_capital_list[len(description_capital_list)-1]
#         WS.range(f"P{i}").value = "More than one"
#     else:
#         if len(description_capital_list) == 0:
#             continue
#         mainWord = description_capital_list[0]
#
#     WS.range(f"O{i}").value = mainWord

coarse_list = ["SAND", "GRAVEL", "GROUT"]
fine_list = ["SILT", "CLAY"]

for i in range(2, 81):
    material_list = []
    currentDepth = float(WS2.range(f'C{i}').value)
    currentDepth_lower = currentDepth + 0.5
    vbhName = WS2.range(f'B{i}').value
    #print(vbhName, float(currentDepth))
    for j in range(2, 66):
        if vbhName != WS.range(f'B{j}').value:
            continue
        if currentDepth >= float(WS.range(f'C{j}').value) and currentDepth < float(WS.range(f'D{j}').value):
            material_list.append(WS.range(f'O{j}').value)
            #index_list.append(j)
        if currentDepth < float(WS.range(f'C{j}').value) and currentDepth_lower >= float(WS.range(f'D{j}').value):
            material_list.append(WS.range(f'O{j}').value)
            #index_list.append(j)
            break
    print(vbhName, float(currentDepth), material_list)
    if material_list[0] in coarse_list:
        WS2.range(f'E{i}').value = "Coarse"
    elif material_list[0] in fine_list:
        WS2.range(f'E{i}').value = "Fine"
    else:
        WS2.range(f'E{i}').value = "Rock"