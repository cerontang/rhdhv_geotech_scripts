import xlwings as xw
from difflib import SequenceMatcher

XL_path = r'C:\Users\921722\Box\BI3753 Team\30 - Geotech\06-Grouting works\02 - Verification\01 Copy of all verification BHs\01 Final Logs to Check\Done (SPT Check)\ERI_Final Logs FGE Check_221214.xlsx'
WB = xw.Book(XL_path)
WS = WB.sheets["ERI_MAIN"]

for i in range (2,178):
    remarks = WS.range(f"H{i}").value
    nature = WS.range(f"I{i}").value
    if remarks != "Change in Final Log":
        continue
    if nature == "Newly added in Final Log":
        continue
    description_final = str(WS.range(f"E{i}").value).strip().replace(",", "").replace(".", "")
    description_draft = str(WS.range(f"G{i}").value).strip().replace(",", "").replace(".", "")
    #print(description_draft, description_final)
    final_list = description_final.split()
    draft_list = description_draft.split()
    original_list = []
    change_list = []
    for item in draft_list:
        if not item.isupper():
            continue
        if item not in description_final:
            original_list.append(item)
    for item in final_list:
        if not item.isupper():
            continue
        if item not in description_draft:
            change_list.append(item)
    print(i, original_list, change_list)
    try:
        print(f"Changed to {change_list[0]} from {original_list[0]}")
        WS.range(f"I{i}").value = f"Changed to {change_list[0]} from {original_list[0]}"
    except:
        pass

