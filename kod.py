import openpyxl

wb_obj = openpyxl.load_workbook("oko_2.xlsx")
sheet_obj = wb_obj.active

"""
cell_obj = sheet_obj.cell(row = 3, column = 2)
kriteri = sheet_obj.cell(row = 3, column = 1)
  pokazatel = sheet_obj.cell(row = i, column = 1)
    print(pokazatel.value)
    print(str(pokazatel.value) + ' - ' )
    i = i+1
print(cell_obj.value) - названия столбцов
print(sheet_obj.max_row) - количество строк
print(sheet_obj.max_column) - количество столбцов
"""
#Привязана динамическая таблица с автоматическим обновлением значений
#m - максимальное количество строк
#c - максимальное количество столбов
#ssred - сумма показателей по строке
#sred - среднее значение
#b - счетчик количества слагаемых
m = sheet_obj.max_row
c = sheet_obj.max_column
pokazatel = 0
m = m+1 
i = 2
j= 2
c = sheet_obj.max_column
ssred = 0
sred=0
b=0
while j < c and i<m:
    znach = sheet_obj.cell(row = i, column = j)
    pokazatel = sheet_obj.cell(row = i, column = 1)
    if znach.value!=None:
        ssred = ssred + znach.value
        j=j+1
        b = b + 1
    else:           
        j = 2
        i = i + 1
        sred = ssred / b
        ssred = 0
        b = 0
        print(str(pokazatel.value) + " " + str(sred))
    
    
    
    
    
'''    
print(pokazatel)
while i < m: 
    pokazatel = sheet_obj.cell(row = i, column = 3)
    kriteri = sheet_obj.cell(row = i, column = 1)
    print(pokazatel.value)
    #print(str(kriteri.value) + ' - ' + str(pokazatel.value))
    i = i+1

"""
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
matplotlib.style.use('ggplot')

#подставьте ссылку на ваш файл или полный путь к файлу на вашем компьютере...    
url = 'https://docs.google.com/spreadsheet/ccc?key=0Ak1ecr7i0wotdGJmTURJRnZLYlV3M2daNTRubTdwTXc&output=csv'

df = pd.read_csv(url, names=['val','date'], index_col=[1], decimal=',',
                 parse_dates=True, dayfirst=True)

df.plot()
#plt.savefig('d:/temp/out.png')
plt.show()
"""

'''