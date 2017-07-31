
import openpyxl as xl
import random

#ファイルの読み込み

wb = xl.load_workbook('WordList.xlsx')

sheet = wb.get_sheet_by_name('Sheet1')

shsheet = wb.get_sheet_by_name('shuffleWord')

#スタート値とエンド値と個数の設定（今後入力型にする予定）

s = 1
e = sheet.max_row
choice = e - s + 1

#乱数生成

for i in range(1,sheet.max_row+1):
    sheet.cell(row=i,column=3).value = random.random()

#乱数から指定した区間の部分をソートして、shuffleWordに格納

a = []

for i in range(s,e+1):
    a = a + [[sheet.cell(row=i,column=1).value,sheet.cell(row=i,column=2).value,sheet.cell(row=i,column=3).value]]

# sorted(object or list,key=lambda object:object.choice)
b = sorted(a,key=lambda a:a[2])

for i in range(choice):
    shsheet.cell(row=i+2,column=1).value = b[i][0]
    shsheet.cell(row=i+2,column=2).value = b[i][1]

#save

wb.save('WordList.xlsx')
