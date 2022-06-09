import openpyxl

wb = openpyxl.load_workbook("Pagi_A.xlsx")
print(wb.sheetnames)
wb.active = 1
sh = wb.active
print(sh)

import numpy as np
a = np.zeros((163), 'U163')
b = np.zeros(163)
for j in range(143, 163):
    a[j]= sh['C' + str(j)].value
    b[j]= sh['L' + str(j)].value
    print(a[j],'=',b[j])

import matplotlib.pyplot as plt
plt.pie(b, labels=a, autopct = '%1.2f%%')
plt.title('Diagram pie kolom volume')
plt.show()

