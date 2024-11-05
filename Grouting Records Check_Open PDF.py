import os
import xlwings as xw

XL_path = r'C:\Users\921722\Box\BI3753 Team\30 - Geotech\03-Level Surveys\Pier 1 Level surveys\20230127\Pier 1 Heave Assessment_230222.xlsx'
WB = xw.Book(XL_path)
WS = WB.sheets["All Heave Areas"]

rowNumber = "247"
probeName = str(WS.range(f"A{rowNumber}").value)
probeUtil = probeName.split("-")[0] + "-" + probeName.split("-")[1]
print(probeUtil)

grouting_records_dir_list = ["PIRE-1_Grouting_Record_PART-01_230315", "PIRE-1_Grouting_Record_PART-02_230315", "PIRE-2_Grouting_Record_230315"]
for dir in grouting_records_dir_list:
    if dir == "PIRE-2_Grouting_Record_230315":
        path = f"C:/Users/921722/Box/BI3753 Team/30 - Geotech/06-Grouting works/01 - Grouting Records/Pier 2/{dir}"
    else:
        path = f"C:/Users/921722/Box/BI3753 Team/30 - Geotech/06-Grouting works/01 - Grouting Records/Pier 1/Individuals/{dir}"
    if not os.path.exists(f"{path}/{probeUtil}"):
        continue
    gp_path = f"{path}/{probeUtil}/{probeName}"
    gp_files = os.listdir(gp_path)
    for file in gp_files:
        if "permeation" in file.lower():
            open_path = f"{gp_path}/{file}"
            os.startfile(open_path)
            print("File Found")

