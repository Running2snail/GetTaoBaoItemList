# -*- coding:UTF-8 -*-

import xlwt

# 创建workbook和sheet对象 注意Workbook的开头W要大写
# workbook = xlwt.Workbook()
# sheet1 = workbook.add_sheet('sheet1', cell_overwrite_ok=True)
# sheet2 = workbook.add_sheet('sheet2', cell_overwrite_ok=True)
#
# # 向sheet写入数据
# sheet1.write(0, 0, 'this')
# sheet1.write(0, 1, 'is')
# sheet2.write(0, 0, 'four')
# sheet2.write(0, 1, 'sheet')
#
# workbook.save('Workbook2.xls')
# print('创建execel完成！')

workbook = xlwt.Workbook()
sheet1 = workbook.add_sheet('sheet1', cell_overwrite_ok=True)
title = ['姓名', '年龄', '性别', '分数']
stus = [['lili', 20, '女', 89.9], ['lucy', 21, '女', 90.9], ['make', 22, '男', 91.9], ['mary', 23, '男', 92.9]]
for i in range(len(title)):
    sheet1.write(0, i, title[i])

for i in range(len(stus)):
    for j in range(4):
        sheet1.write(i + 1, j, stus[i][j])

workbook.save('Workbook2.xls')
print('创建execel完成！')