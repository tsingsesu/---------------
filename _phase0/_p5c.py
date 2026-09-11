import openpyxl, os
W = r"C:\Users\lhl87\Desktop\高教社杯全国大学生数学建模大赛"
wb = openpyxl.load_workbook(os.path.join(W,"附件","附件5","result2.xlsx"))
ws = wb["计划购电量"]
hdr = [c.value for c in ws[1]]
print("COUNTS:", len(hdr))
print("FIRST10:", hdr[:11])
print("LAST5:", hdr[-5:])
for nm in ["充放电量","紧急购电量"]:
    ws2 = wb[nm]
    print("===", nm, ws2.max_row, ws2.max_column)
    for r in ws2.iter_rows(values_only=True):
        print("   ", r)
