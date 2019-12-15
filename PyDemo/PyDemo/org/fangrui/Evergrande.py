import xlrd


def excelProcess():
    excel = xlrd.Workbook(encoding='utf-8')
    sheet = excel.add_sheet("123", cell_overwrite_ok=False)
    temp_value = 'HYPERLINK("https://wwww.baidu.com";"百度一下")'
    sheet.write(1, 1, xlrd.Formula(temp_value))
    excel.save("1231.xls")


excelProcess()
