import openpyxl
wb=openpyxl.load_workbook("Pagi_A.xlsx")
print(wb.sheetnames)
wb.active=3
sh=wb.active
print(sh)
import numpy as np
x=np.zeros((143), 'U143')
y=np.zeros(143)

for i in range(123,143):
    x[i] =sh['B' + str(i)].value
    y[i] =sh['L' + str(i)].value
    print(x[i],y[i])

import matplotlib.pyplot as plt
plt.bar(x, y, label="Blue Bar", color='b')
plt.plot()
plt.xlabel("Data volume 121 - 140", size=16)
plt.ylabel("Data per meter kubik 121 - 140", size=14)
plt.title("Chart Volume", size=18 )
plt.legend()
plt.show()