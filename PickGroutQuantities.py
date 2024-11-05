import xlwings as xw

source_path = r"C:\Users\921722\OneDrive - Royal HaskoningDHV\20220902AD\VBH_Grout Information_20220816_check_2.xlsx"
WB = xw.Book(source_path)
WS = WB.sheets['FGE_FUGRO_2']
WS2 = WB.sheets['GROUT_FUGRO_2 0.5']
#index_list = []

for i in range(906, 1969):
    groutList_inrange = []
    currentDepth = float(WS2.range(f'C{i}').value)
    currentDepth_lower = currentDepth + 0.5
    vbhName = WS2.range(f'B{i}').value
    #print(vbhName, float(currentDepth))
    for j in range(2, 1002):
        # if j in index_list:
        #     continue
        if vbhName != WS.range(f'B{j}').value:
            continue
        if currentDepth >= float(WS.range(f'C{j}').value) and currentDepth < float(WS.range(f'D{j}').value):
            groutList_inrange.append(WS.range(f'G{j}').value)
            #index_list.append(j)
        if currentDepth < float(WS.range(f'C{j}').value) and currentDepth_lower >= float(WS.range(f'D{j}').value):
            groutList_inrange.append(WS.range(f'G{j}').value)
            #index_list.append(j)
            break
    if len(groutList_inrange) == 0:
        groutList_inrange.append(float(0.00))
    grout_avg = sum(groutList_inrange)/len(groutList_inrange)
    #print(groutList_inrange)
    print(vbhName, float(currentDepth), grout_avg, groutList_inrange)
    WS2.range(f'D{i}').value = grout_avg