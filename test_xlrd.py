# -*- coding:UTF-8 -*-

import xlrd

workbook = xlrd.open_workbook('Workbook1.xlsx')
print(workbook.sheet_names())
workbook1 = workbook.sheet_by_name(u'Sheet1')
num_rows = workbook1.nrows
# for curr_row in range(num_rows):
#     row_value = workbook1.row_values(curr_row)
#     print('row%s value is %s' %(curr_row, row_value))

num_cols = workbook1.ncols
# for curr_col in range(num_cols):
#     col_value = workbook1.col_values(curr_col)
#     print('col%s value is %s' % (curr_col, col_value))

for row in range(num_rows):
    for col in range(num_cols):
        cell = workbook1.cell_value(row, col)
        print(cell)