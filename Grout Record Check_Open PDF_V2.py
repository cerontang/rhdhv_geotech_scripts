import os

probeName = str(input())
record_type = ["permeation", "drilling", "block"]
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
        if record_type[0] in file.lower():
            open_path = f"{gp_path}/{file}"
            os.startfile(open_path)
            print("File Found")