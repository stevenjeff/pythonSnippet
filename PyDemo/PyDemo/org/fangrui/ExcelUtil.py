import xlrd

file_path = r'C:\Users\zhangfangrui\Desktop\需求文档\机票\机票验舱验价.xlsx'
workbook = xlrd.open_workbook(file_path)  # 打开excel工作簿
sheet = workbook.sheet_by_index(0)  # 选择第一张sheet
for row in range(sheet.nrows):  # 第一个for循环遍历所有行
    print()
    # for col in range(sheet.ncols):  # 第二个for循环遍历所有列，这样就找到某一个xy对应的元素，就可以打印出来
    #     print("%s" % sheet.row(row)[col].value, '\t', end='')
    print(" /*")
    print("  * " + sheet.row(row)[2].value)
    print("  */")
    print("@JacksonXmlProperty(localName = \"" + sheet.row(row)[1].value + "\")")
    print("private %s" % sheet.row(row)[0].value, '', end='')
    print("%s" % sheet.row(row)[1].value, '', end='')
