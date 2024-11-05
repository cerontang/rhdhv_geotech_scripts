import xlwings as xw

file_path = "C:/Users/921722/OneDrive - Royal HaskoningDHV/20220713 AD/STORM WATER MANHOLES ( INVERT  COVER LEVEL)_ LANDSIDE 07-07-2022.xlsx"
WB = xw.Book(file_path)
WS = WB.sheets["Check"]

# for i in range(2, 334):
#     if WS.range(f'B{i}').value == "MH IL":
#         if WS.range(f'F{i}').value is None:
#             for j in range (0, 10):
#                 if WS.range(f'H{i+j}').value != "MH. IL":
#                     continue
#                 if WS.range(f'H{i+j}').value == "MH. IL":
#                     WS.range(f'F{i}').value = WS.range(f'L{i+j}').value
#                     break

for i in range(2, 334):
    if WS.range(f'I{i}').value is not None or WS.range(f'J{i}').value is not None:
        WS.range(f'N{i}').value = WS.range(f'H{i}').value
        WS.range(f'O{i}').value = WS.range(f'I{i}').value
        WS.range(f'P{i}').value = WS.range(f'J{i}').value
        for j in range (0, 10):
            if WS.range(f'H{i+j}').value == 'MH. IL':
                WS.range(f'N{i+1}').value = WS.range(f'H{i+j}').value
                WS.range(f'Q{i+1}').value = WS.range(f'K{i+j}').value
                WS.range(f'R{i+1}').value = WS.range(f'L{i+j}').value
                break
        continue
    if WS.range(f'H{i}').value == 'MH. IL':
        continue
    WS.range(f'N{i+1}').value = WS.range(f'H{i}').value
    WS.range(f'Q{i+1}').value = WS.range(f'K{i}').value
    WS.range(f'R{i+1}').value = WS.range(f'L{i}').value
