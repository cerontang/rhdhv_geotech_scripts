import os
import pandas as pd

VBH_directory = r"C:\Users\921722\Box\BI3753 Team\30 - Geotech\06-Grouting works\02 - Verification\01 Copy of all verification BHs\Pier 3"
VBH_list = os.listdir(VBH_directory)
master_VBH_list = []
master_VBH_beautified_list = []
master_link_list = []

for VBH in VBH_list:
    if not (VBH.startswith("A-") or VBH.startswith("R-") or VBH.startswith("VBH-")):
        if not (VBH.startswith("Inundation") or VBH.startswith("Checked")):
            continue
        VBH_subdirectory = f"{VBH_directory}\{VBH}"
        VBH_sublist = os.listdir(VBH_subdirectory)
        for VBH_sub in VBH_sublist:
            if not (VBH_sub.startswith("A-") or VBH_sub.startswith("R-") or VBH_sub.startswith("VBH-")):
                continue
            VBH_subpath = f"{VBH_subdirectory}\{VBH_sub}"
            print(VBH_subpath)
            master_VBH_list.append(VBH_sub)
            master_link_list.append(VBH_subpath)

        continue

    VBH_path = f"{VBH_directory}\{VBH}"
    print(VBH_path)
    master_VBH_list.append(VBH)
    master_link_list.append(VBH_path)

for vbh in master_VBH_list:
    vbh = vbh.replace(".pdf", "")
    vbh = vbh.replace("_F2", "")
    vbh = vbh.replace("_F", "")
    vbh = vbh.replace("_1", "")
    master_VBH_beautified_list.append(vbh)

output_list = list(zip(master_VBH_beautified_list, master_link_list))
df = pd.DataFrame(output_list)
df.to_csv(r"csv\P3 missing_VBH_link_list.csv")