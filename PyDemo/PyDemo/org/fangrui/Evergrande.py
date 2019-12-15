import xlwings as xw
import time
import shutil
import os

excelFileName = r'H:\新建文件夹\3、建安-新 - 副本 - 副本.xlsx'
companyPath = "H:\\新建文件夹\\3.建安图片"
linkColumn = 22
totalSheetName = "汇总"
detailsheetName = "合并明细（带合同号）"
range = "A3:Y652"
ifVisible = False
voucherDir = "凭证"
notContractvoucherDir = "无合同"

def excelProccessXlwings():
    print("begin:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))

    app = xw.App(visible=ifVisible, add_book=False)
    # 连接到excel
    workbook = app.books.open(excelFileName)  # 连接excel文件
    # 保存
    totalSheet = workbook.sheets[totalSheetName]
    detailSheet = workbook.sheets[detailsheetName]
    rng = totalSheet.range(range).value
    for row in rng:
        print(row[3] + "--" + row[2] + "--" + str(row[13]))
    # workbook.save(".\\3、建安-新 - 副本 - 副本_new1.xlsx")
    workbook.close()
    app.quit()
    print("end:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))


def walkDir(dir):
    for home, dirs, files in os.walk(dir, followlinks=False, topdown=True):
        print(files)
        # print(dirs)
        # print(files)
        # for dir in dirs:
        #     print(dir)


walkDir(companyPath)
# excelProccessXlwings()
