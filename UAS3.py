import openpyxl

wb = openpyxl.load_workbook("Pagi_A.xlsx")
print(wb.sheetnames)
wb.active = 2
sh = wb.active
print(sh)

for  i in range (103, 123):
    a = sh['E' + str(i)].value
    b = sh['J' + str(i)].value
    c = sh['K' + str(i)].value

    print('sebelumnya =',a)
    print('penutupan =',b)
    print('selisihnya =',c)
    print()