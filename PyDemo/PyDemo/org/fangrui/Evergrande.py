import xlwings as xw
import time
import shutil
import os

rootPath = "D:\\360Downloads\\evergrande\\"
excelName = "3、建安-新 - 副本 - 副本.xlsx"
excelNameReName = "3、建安-新 - 副本 - 副本_done.xlsx"
projectName = "3.建安图片"
projectRoot = rootPath + projectName
linkColumn = 22
labelColumn = 24
valueColumn = "L"
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
    workbook = app.books.open(rootPath + excelName)  # 连接excel文件
    # 保存
    totalSheet = workbook.sheets[totalSheetName]
    detailSheet = workbook.sheets[detailsheetName]
    rng = totalSheet.range(range)
    currentCompanyIndex = 0
    for row in rng.rows:
        if row[0].value is not None:
            currentCompanyIndex = int(row[0].value)
        if row[3].value is None or row[2].value is None or row[11].value is None or row[3].value == "无合同" or row[
            2].value == "无合同" or str(
                row[3].value).strip() == "" or str(row[2].value).strip() == "":
            row[labelColumn].value = "无合同"
            continue
        companyNameDir = getCompanyDir(currentCompanyIndex)
        if companyNameDir is None:
            continue
        currentCompanyDir = projectRoot + "\\" + companyNameDir
        row[linkColumn].formula = '=HYPERLINK(".\\' + projectName + '\\' + companyNameDir + '",' + valueColumn + str(
            row.row) + ')'
        print(getCompanyDir(currentCompanyIndex))
        print(row[3].value + "--" + row[2].value + "--" + str(row[11].value))
    workbook.save(rootPath + excelNameReName)
    workbook.close()
    app.quit()
    print("end:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))


def walkDir():
    folderlist = os.listdir(projectRoot)  # 列举文件夹
    for currentDir in folderlist:
        companyDir = projectRoot + "\\" + currentDir
        if os.path.isdir(companyDir):
            for company in os.listdir(companyDir):
                print(companyDir + "\\" + company)

def getCompanyDir(currentCompanyIndex):
    folderlist = os.listdir(projectRoot)  # 列举文件夹
    for currentDir in folderlist:
        companyDir = projectRoot + "\\" + currentDir
        dirs = currentDir.split("、")
        if os.path.isdir(companyDir) and str(currentCompanyIndex) == dirs[0]:
            return currentDir;


# walkDir()
excelProccessXlwings()

# if "[" in company or "]" in company or "【" in company or "】" in company:
#     print(companyDir +"\\"+company)
# shutil.rmtree(companyDir + "\\" + company)
# for home, dirs, files in os.walk(dir, followlinks=False, topdown=True):
# print(files)
# print(dirs)
# print(files)
# for dir in dirs:
#     print(dir)
