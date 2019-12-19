import os
import shutil
import time

import xlwings as xw

rootPath = "D:\\360Downloads\\evergrande\\"
excelName = "3、建安-新 - 副本 - 副本.xlsx"
excelNameReName = "3、建安-新 - 副本 - 副本_done.xlsx"
projectName = "3.建安图片"
# 汇总excel范围
sumRange = "A3:Y652"
# 合并明细excel范围
detailRange = "A2:U2381"
projectRoot = rootPath + projectName
linkColumn = 22
labelColumn = 24
valueColumn = "L"
contractNoIndex = 3
# 明细合同列
contractNoIndexDetail = 22
# 明细年
yearIndexDetail = 0
# 明细月
monthIndexDetail = 1
# 凭证号
voucherIndexdetail = 2
totalSheetName = "汇总"
detailsheetName = "合并明细（带合同号）"
ifVisible = False
voucherDirName = "凭证"
notContractvoucherDirName = "无合同"

def excelProccessXlwings():
    print("begin:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))

    app = xw.App(visible=ifVisible, add_book=False)
    # 连接到excel
    workbook = app.books.open(rootPath + excelName)  # 连接excel文件
    # 保存
    totalSheet = workbook.sheets[totalSheetName]
    detailSheet = workbook.sheets[detailsheetName]
    rng = totalSheet.range(sumRange)
    currentCompanyIndex = 0
    for row in rng.rows:
        if row[0].value is not None:
            currentCompanyIndex = int(row[0].value)
        if row[contractNoIndex].value is None or row[2].value is None or row[11].value is None or row[
            contractNoIndex].value == "无合同" or row[
            2].value == "无合同" or str(
            row[contractNoIndex].value).strip() == "" or str(row[2].value).strip() == "":
            row[labelColumn].value = "无合同"
            continue
        companyNameDir = getCompanyDir(currentCompanyIndex)
        if companyNameDir is None:
            continue
        contractMoneyDir = row[3].value + "--" + row[2].value + "--" + str(row[11].value)
        copyFiles(companyNameDir, row[contractNoIndex], contractMoneyDir, detailSheet)
        row[linkColumn].formula = '=HYPERLINK(".\\' + projectName + '\\' + companyNameDir + '",' + valueColumn + str(
            row.row) + ')'
        print(getCompanyDir(currentCompanyIndex))
        print(row[3].value + "--" + row[2].value + "--" + str(row[11].value))
    workbook.save(rootPath + excelNameReName)
    workbook.close()
    app.quit()
    print("end:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))


def copyFiles(companyNameDir, contractNo, contractMoneyDir, detailSheet):
    currentCompanyDir = projectRoot + "\\" + companyNameDir + "\\"
    contractDir = currentCompanyDir + companyNameDir.split("、")[1].strip() + "\\"
    contractMoneyDirFullPath = projectRoot + "\\" + companyNameDir + "\\" + contractMoneyDir
    if not os.path.exists(contractMoneyDirFullPath):
        os.mkdir(contractMoneyDirFullPath)
    contractfolderlist = os.listdir(contractDir)  # 列举文件夹
    # 拷贝合同
    for currentDir in contractfolderlist:
        dircontractName = contractNoReplace(currentDir)
        contractName = contractNoReplace(contractNo)
        if contractName in dircontractName:
            shutil.copytree(contractDir + currentDir, contractMoneyDirFullPath)
    # 拷贝凭证
    rng = detailSheet.range(detailRange)
    for row in rng.rows:
        detailContractNo = row[contractNoIndexDetail]
        if detailContractNo is None or detailContractNo == '':
            proccessNoContractVoucher()
            continue
        copyVoucher(detailSheet, detailContractNo, contractMoneyDirFullPath)


def copyVoucher(detailSheet, contractNo, contractMoneyDirFullPath):
    rng = detailSheet.range(detailRange)
    for row in rng.rows:
        detailContractNo = row[contractNoIndexDetail]
        if detailContractNo is None or detailContractNo == '':
            proccessNoContractVoucher()
            continue
        if contractNo == detailContractNo:
            year = row[yearIndexDetail]
            month = row[monthIndexDetail]
            voucherStrNo = row[voucherIndexdetail]
            foundVoucherDir = getMatchVoucherDir(int(year), int(month), voucherStrNo, currentCompanyDir)
            if foundVoucherDir == "":
                continue
            shutil.copytree(foundVoucherDir, contractMoneyDirFullPath)


def getMatchVoucherDir(year, month, voucherStrNo, currentCompanyDir):
    voucherStrArray = voucherStrNo.split("-");
    voucherNumStr = voucherStrArray[1]
    voucherNo = int(voucherNumStr)
    voucherDir = currentCompanyDir + voucherDirName + "\\"
    if not os.path.exists(voucherDir):
        return ""
    voucherfolderlist = os.listdir(voucherDir)  # 列举文件夹
    for currentDir in voucherfolderlist:
        voucherDirArray = currentDir.split("-")
        voucherYearFromDir = int(voucherDirArray[0])
        voucherMonthFromDir = int(voucherDirArray[1])
        voucherNoFromDir = int(voucherDirArray[2])
        if year == voucherYearFromDir and month == voucherMonthFromDir and voucherNo == voucherNoFromDir:
            return voucherDir + currentDir
    return ""


def voucherCompare(year, month, voucherNo, currentDir):
    possibleVoucherDirName = ""
    possibleVoucherDirName = str(year) + "年" + str(month) + "月"


def proccessNoContractVoucher():
    return

def contractNoReplace(contractNo):
    return contractNo.replace("[", "").replace("]", "").replace("【", "").replace("】", "")

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
