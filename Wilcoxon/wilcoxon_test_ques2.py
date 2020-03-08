#!/bin/python3

from itertools import combinations
from wilcoxon_test import cal_wilcoxon, open_sheet


def cal_avg(iterable):
    """Function to calculate average of products between customers"""
    data = []
    sheet_name = 'W'
    for i, item in enumerate(iterable):
        indx = nonkey_cols.index(item)
        data.append(list([sheet.cell_value(r, indx+2) for r in range(1, sheet.nrows)]))
        sheet_name = sheet_name + '_' + str(item)
    return sheet_name, [sum(item)/len(item) for item in zip(*data)]


if __name__ == '__main__':

    sheet = open_sheet()

    key = [sheet.cell_value(r, 1) for r in range(1, sheet.nrows)]  # taking key for 48806
    nonkey_cols = list([int(cell.value) for cell in sheet.row(0)[2:]])   # taking all non-key columns

    com = []
    """Finding all possible combinations"""
    for i in range(1, len(nonkey_cols)):
        com.append(list(combinations(nonkey_cols, i+1)))

    """Running Wilcoxon test for all combinations found"""
    for row in com:
        for i in row:
            sheet_n, nonkey = cal_avg(i)
            print('Value of W for ' + str(sheet_n) + ' is', str(int(cal_wilcoxon(key, nonkey))))
