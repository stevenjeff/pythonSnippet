''ffff'''
import os
import shutil
import time

import xlwings as xw

# 根目录 ，包含excel的目录
rootPath = "D:\\360Downloads\\evergrande\\"
# excel 源文件名称
excelName = "3、建安-新 - 副本 - 副本.xlsx"
# 保存重命名
excelNameReName = "3、建安-新 - 副本 - 副本_done.xlsx"
# 被操作文件夹
projectName = "3.建安图片"
# projectRoot
projectRoot = rootPath + projectName
# excel 操作是否可见
ifVisible = False

# 汇总配置
# 汇总sheet 配置
totalSheetName = "汇总"
# 汇总excel范围
sumRange = "A3:Y652"
# excel 中 数字 文件夹链接所在列
linkColumn = 22
# 标记无合同 所在列
labelColumn = 24
# 取账面成本所在列 字母号
valueColumn = "L"
# 索引列
serial_no_column_index = 0
# 合同编号所在列
contractNoIndex = 3
# 合同名所在列
contract_name_index = 2
# 账面进成本列
carrying_cost_index = 11

# 明细sheet配置
detailsheetName = "合并明细（带合同号）"
# 合并明细excel范围
detailRange = "A2:U2381"
# 明细合同列
contractNoIndexDetail = 20
# 明细年
yearIndexDetail = 0
# 明细月
monthIndexDetail = 1
# 凭证号
voucherIndexdetail = 3
# 凭证文件夹名称
voucherDirName = "凭证"
# 没有合同对应的凭证所在文件夹
notContractvoucherDirName = "无合同"
# 合同文件夹可选名称
optionalContractFolderName = "合同"


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
    voucherCopiedInstance = None
    for row in rng.rows:
        if row[serial_no_column_index].value is not None:
            currentCompanyIndex = int(row[serial_no_column_index].value)
            proccessNoContractVoucherFolder(voucherCopiedInstance)
            voucherCopiedInstance = VoucherCopied(set(), "", "")
        companyNameDir = getCompanyDir(currentCompanyIndex)
        if companyNameDir is None:
            continue
        if row[contractNoIndex].value is None or row[contract_name_index].value is None or row[
            carrying_cost_index].value is None or row[
            contractNoIndex].value == "无合同" or row[
            contract_name_index].value == "无合同" or str(
            row[contractNoIndex].value).strip() == "" or str(row[contract_name_index].value).strip() == "":
            row[labelColumn].value = "无合同"
            row.color = (255, 0, 0)
            row[
                linkColumn].formula = '=HYPERLINK("..\\' + projectName + '\\' + companyNameDir + '\\' + '",' + valueColumn + str(
                row.row) + ')'
            continue
        contractMoneyDir = row[contractNoIndex].value + "--" + row[contract_name_index].value + "--" + str(
            row[carrying_cost_index].value)
        copyFiles(companyNameDir, row[contractNoIndex].value, contractMoneyDir, detailSheet, voucherCopiedInstance)
        row[
            linkColumn].formula = '=HYPERLINK("..\\' + projectName + '\\' + companyNameDir + '\\' + contractMoneyDir + '",' + valueColumn + str(
            row.row) + ')'
        # 写一行存一次不知道性能是否 影响很大
        # workbook.save(rootPath + excelNameReName)
        # print(getCompanyDir(currentCompanyIndex))
        # print(row[3].value + "--" + row[2].value + "--" + str(row[11].value))
    workbook.save(rootPath + excelNameReName)
    workbook.close()
    app.quit()
    print("end:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))


def proccessNoContractVoucherFolder(voucherCopied):
    if voucherCopied is not None and voucherCopied.voucherDir != "" and voucherCopied.companyNameDir != "":
        vouchers = os.listdir(voucherCopied.voucherDir)
        for voucherDirname in vouchers:
            if voucherDirname not in voucherCopied.copiedVoucherFolderNames:
                if not os.path.exists(voucherCopied.companyNameDir + notContractvoucherDirName):
                    os.mkdir(voucherCopied.companyNameDir + notContractvoucherDirName)
                if not os.path.exists(voucherCopied.companyNameDir + notContractvoucherDirName + "\\" + voucherDirname):
                    shutil.copytree(voucherCopied.voucherDir + voucherDirname,
                                    voucherCopied.companyNameDir + notContractvoucherDirName + "\\" + voucherDirname)


def copyFiles(companyNameDir, contractNo, contractMoneyDir, detailSheet, voucherCopiedInstance):
    currentCompanyDir = projectRoot + "\\" + companyNameDir + "\\"
    contractDir = currentCompanyDir + companyNameDir.split("、")[1].strip() + "\\"
    optionalContractDir = currentCompanyDir + optionalContractFolderName + "\\"
    voucherDir = currentCompanyDir + voucherDirName + "\\"
    voucherCopiedInstance.voucherDir = voucherDir
    voucherCopiedInstance.companyNameDir = currentCompanyDir
    contractMoneyDirFullPath = projectRoot + "\\" + companyNameDir + "\\" + contractMoneyDir
    if not os.path.exists(contractMoneyDirFullPath):
        os.mkdir(contractMoneyDirFullPath)
    # 拷贝合同
    print("拷贝合同 begin:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
    copyContract(contractDir, optionalContractDir, contractNo, contractMoneyDirFullPath, detailSheet, voucherDir,
                 voucherCopiedInstance)
    print("拷贝合同 end:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))


def copyContract(contractDir, optionalContractDir, contractNo, contractMoneyDirFullPath, detailSheet, voucherDir,
                 voucherCopiedInstance):
    actual_contract_dir = ""
    if os.path.exists(contractDir):
        actual_contract_dir = contractDir
    elif os.path.exists(optionalContractDir):
        actual_contract_dir = optionalContractDir
    else:
        return
    contractfolderlist = os.listdir(actual_contract_dir)  # 列举文件夹
    # 拷贝合同
    for currentDir in contractfolderlist:
        dircontractName = contractNoReplace(currentDir)
        contractName = contractNoReplace(contractNo)
        if contractName in dircontractName:
            if os.path.isdir(actual_contract_dir + currentDir):
                if not os.path.exists(contractMoneyDirFullPath + "\\" + currentDir):
                    shutil.copytree(actual_contract_dir + currentDir, contractMoneyDirFullPath + "\\" + currentDir)
            else:
                if not os.path.exists(contractMoneyDirFullPath + "\\" + currentDir):
                    shutil.copy(actual_contract_dir + currentDir, contractMoneyDirFullPath)
                # 拷贝凭证
            print("拷贝凭证 begin:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
            copyVoucher(detailSheet, contractMoneyDirFullPath, voucherDir, contractName, voucherCopiedInstance)
            print("拷贝凭证 end:", time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))


def copyVoucher(detailSheet, contractMoneyDirFullPath, voucherDir, contractNameFromTotal, voucherCopiedInstance):
    rng = detailSheet.range(detailRange)
    for row in rng.rows:
        detailContractNo = row[contractNoIndexDetail].value
        print("凭证对应合同号:", detailContractNo)
        if detailContractNo is None or detailContractNo == '':
            continue
        formatdetailContractNo = contractNoReplace(detailContractNo);
        formatcontractNameFromTotal = contractNoReplace(contractNameFromTotal);
        if formatdetailContractNo == formatcontractNameFromTotal:
            copyVoucherInternal(row, contractMoneyDirFullPath, voucherDir, voucherCopiedInstance)


def copyVoucherInternal(row, contractMoneyDirFullPath, voucherDir, voucherCopiedInstance):
    year = row[yearIndexDetail].value
    month = row[monthIndexDetail].value
    voucherStrNo = row[voucherIndexdetail].value
    foundVoucherDir = getMatchVoucherDir(int(year), int(month), voucherStrNo, voucherDir)
    if foundVoucherDir == "":
        return
    print("拷贝凭证号:", voucherStrNo, " 凭证路径：", voucherDir + foundVoucherDir)
    voucherCopiedInstance.copiedVoucherFolderNames.add(foundVoucherDir)
    if os.path.exists(contractMoneyDirFullPath + "\\" + foundVoucherDir):
        return
    shutil.copytree(voucherDir + foundVoucherDir, contractMoneyDirFullPath + "\\" + foundVoucherDir)


def getMatchVoucherDir(year, month, voucherStrNo, voucherDir):
    voucherStrArray = voucherStrNo.split("-");
    voucherNumStr = voucherStrArray[1]
    voucherNo = int(voucherNumStr)
    if not os.path.exists(voucherDir):
        return ""
    voucherfolderlist = os.listdir(voucherDir)  # 列举文件夹
    for currentDir in voucherfolderlist:
        voucherDirArray = currentDir.split("-")
        voucherYearFromDir = int(voucherDirArray[0])
        voucherMonthFromDir = int(voucherDirArray[1])
        voucherNoFromDir = int(voucherDirArray[2])
        if year == voucherYearFromDir and month == voucherMonthFromDir and voucherNo == voucherNoFromDir:
            return currentDir
    return ""

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


class VoucherCopied:
    '所有员工的基类'

    def __init__(self, copiedVoucherFolderNames, voucherDir, companyNameDir):
        self.copiedVoucherFolderNames = copiedVoucherFolderNames
        self.voucherDir = voucherDir
        self.companyNameDir = companyNameDir

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
