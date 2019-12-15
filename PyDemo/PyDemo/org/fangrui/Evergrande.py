import xlwings as xw

fileName = r'H:\新建文件夹\3、建安-新 - 副本 - 副本.xlsx'
linkColumn = 22
totalSheetName = "汇总"
detailsheetName = "合并明细（带合同号）"
range = "A3:Y652"

def excelProccessXlwings():
    app = xw.App(visible=True, add_book=False)
    # 连接到excel
    workbook = app.books.open(fileName)  # 连接excel文件
    # 保存
    totalSheet = workbook.sheets[totalSheetName]
    detailSheet = workbook.sheets[detailsheetName]
    rng = totalSheet.range(range).value
    print(rng)
    # workbook.save(".\\3、建安-新 - 副本 - 副本_new1.xlsx")
    workbook.close()
    app.quit()



excelProccessXlwings()
