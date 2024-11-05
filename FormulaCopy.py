
for i in range (5, 60):
    formula = f"=SERIES('Landside Sewer Master R3'!$C${i},'Landside Sewer Master R3'!$BF$3:$CQ$3,'Landside Sewer Master R3'!$BF${i}:$CQ${i},1)"
    print(formula)