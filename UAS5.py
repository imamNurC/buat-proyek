#Coded By Imamnurcakra
import openpyxl
import numpy as np
import matplotlib.pyplot as plt

wb = openpyxl.load_workbook("Pagi_A.xlsx")
print(wb.sheetnames)

x = np.zeros((203), 'U203')
y = np.zeros(203)
for i in wb.worksheets: #mengulang sebanyak banyaknya  lembar kerja yang ada dalam file excel
    if i == wb['02 Feb 2022']: #menyeleksi lembar yang ingin di tentukan yaitu lembar 0,4,6
        wb.active = wb['02 Feb 2022']
        sh = wb.active
        print('sheet active =',sh)
        for j in range(183, 203):
            x[j] = sh['J' + str(j)].value
            y[j] = sh['j' + str(j)].value
            print(y[j],'=',x[j])
        plt.plot(x,y, color='green', marker='o')
        plt.title(i, fontsize=14)
        plt.xlabel('Biaya Kantor', fontsize=14)
        plt.ylabel('Pengeluaran (Rp.)', fontsize=14)
        plt.grid(True)
        plt.show()

    elif i == wb['08 Feb 2022']:
        print('===================')
        wb.active = wb['08 Feb 2022']
        sh = wb.active
        print('sheet active =',sh)
        for k in range(183, 203):
            x[k] = sh['J' + str(k)].value
            y[k] = sh['J' + str(k)].value
            print(y[k],'=',x[k])
        plt.plot(x,y, color='green', marker='o')
        plt.title(i, fontsize=14)
        plt.xlabel('Biaya Kantor', fontsize=14)
        plt.ylabel('Pengeluaran (Rp.)', fontsize=14)
        plt.grid(True)
        plt.show()

    elif i == wb['10 Feb 2022']:
        print('=============')
        wb.active = wb['10 Feb 2022']
        sh = wb.active
        print('sheet active =',sh)
        for l in range(183, 203):
            x[l] = sh['J' + str(l)].value
            y[l] = sh['J' + str(l)].value
            print(y[l], '=', x[l])
        plt.plot(x,y, color='green', marker='o')
        plt.title(i, fontsize=14)
        plt.xlabel('Biaya Kantor', fontsize=14)
        plt.ylabel('Pengeluaran (Rp.)', fontsize=14)
        plt.grid(True)
        plt.show()

