import openpyxl

wb = openpyxl.load_workbook("Pagi_A.xlsx")
print(wb.sheetnames)
wb.active = 0
sh = wb.active
print('sheet active =',sh)

for i in range(163, 183):
    x1 = (sh['H' + str(i)].value + sh['I' + str(i)].value + sh['J' + str(i)].value)/3
    x2 = sh['C' + str(i)].value
    x3 = sh['A' + str(i)].value
    print(x3,x2)
    print("    rata-rata =",x1)
    print()