import xlrd
from datetime import time


def is_merged(row, column):
    for cell_range in sheet.merged_cells:
        row_low, row_high, column_low, column_high = cell_range
        if row in range(row_low, row_high) and column in range(column_low, column_high):
            return row_low, column_low
    return False


rb = xlrd.open_workbook('D:\\project\\ParserLeti\\3_FKTI_osen_2017-2.xls', formatting_info=True)
sheet = rb.sheet_by_index(8)
groups_col = {}
time_of_lessons = {'8:00': None, '9:50': None, '11:40': None, '13:45': None, '15:35': None, '17:25': None}
odd_week = {}
an_even_week = {}
rlo = 0
merged = []
for rownum in range(sheet.nrows):
    c = False
    row = sheet.row_values(rownum)

    if '№ гр.' in row:
        for x in range(len(row)):
            if not isinstance(row[x], str):
                groups_col.setdefault(row[x], x)
                odd_week.setdefault(row[x], None)
                an_even_week.setdefault(row[x], None)
    c = is_merged(rownum, 14)
    if c:
        x, y = c
        val = sheet.cell(x, y)
        fm = rb.xf_list[val.xf_index]
        print(val.value+' is '+ str(fm.dump()))

        if not val.value in merged and val.value != '':
            merged.append(sheet.cell(x, y).value)

print(merged)
