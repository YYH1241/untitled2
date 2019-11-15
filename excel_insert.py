import xlrd
import xlwt
import logging


#TODO:loging.log()
get_data = xlrd.open_workbook("E:\\nlp自动化\\NLPresult.xls")
workbook = xlwt.Workbook(encoding='utf-8')
worksheet = workbook.add_sheet('意图识别结果')
#获取第一张sheet表的内容
table_index = get_data.sheet_by_index(0)
col = table_index.nrows
list = []
i = 1
try:
    for c in range(col):
        if i < col:
            # 获取单行内容
            data = table_index.row_values(i)
            print(data)
            print(data[0])
            for data_index in data:
            # 过滤掉内容为空的部分
                if data_index is '':
                    pass
                else:
                    list.append(data_index)
            i += 1
        else:
            print("excel内容输出完成")
            print(list)
finally:
    print("当前列表长度是"+str(i))

m = 0
n = 0
try:
    for list_index in list:
        data_final = list_index.split(',')
        for insert in data_final:
            print(insert)
            worksheet.write(m, n, insert)
            n += 1
            if n == 6:
                n = 0
                m += 1
            else:
                pass
    else:
        pass
finally:
    print("准备数据完成")

workbook.save("rep_result.xls")

