from FuckExcel import getFuckExcel

fuck_excel = getFuckExcel('./dist/A.xlsx', with_numba=False)
fuck_excel[5:10, 5:10] = 'init'
print(fuck_excel[5:10, 5:10])
fuck_excel.save()
