import xlwings as xw

file_path = r"C:\Users\921722\OneDrive - Royal HaskoningDHV\20220817 AD\Grouting Records\VBH_groutpoints.xlsx"
WB = xw.Book(file_path)
WS = WB.sheets["VBH_groutpoints_rounded"]

def Test_1 ():
    for i in range (3, 357):
        equi_d_count = 0
        gp_d_list = [WS.range(f"D{i}").value, WS.range(f"F{i}").value, WS.range(f"H{i}").value, WS.range(f"J{i}").value]
        for j in range (0,4):
            for k in range (0,4):
                if abs(gp_d_list[j] - gp_d_list[k]) < 0.2:
                    equi_d_count += 1
        print(WS.range(f"B{i}").value, equi_d_count)
        if equi_d_count >= 6 and equi_d_count <= 8:
            WS.range(f"N{i}").value = "Pass"
        else:
            WS.range(f"N{i}").value = "Check"



Test_1()
