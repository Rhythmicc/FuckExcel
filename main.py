from FuckExcel import FuckExcel

fuck_excel = FuckExcel('./A.xlsx')
fuck_excel[5:10, 5:10] = 'init'
print(fuck_excel[5:10, 5:10])
fuck_excel.save()
