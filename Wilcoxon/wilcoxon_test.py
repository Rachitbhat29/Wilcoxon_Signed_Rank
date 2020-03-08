#!/bin/python3

import xlrd
import os

def f_rank(iterable, start=1):
    """Fractional ranking"""
    last, fifo = None, []
    for n, item in enumerate(iterable, start):
        if item[0] != last:
            if fifo:
                mean = sum(f[0] for f in fifo) / len(fifo)
                while fifo:
                    yield mean, fifo.pop(0)[1]
        last = item[0]
        fifo.append((n, item))
    if fifo:
        mean = sum(f[0] for f in fifo) / len(fifo)
        while fifo:
            yield mean, fifo.pop(0)[1]


def cal_wilcoxon(key, nonkey):
    """Function to calculate Wilcoxon Signed Text Statics for each pair of products"""

    diff = [(abs(i-j), (i-j), i, j) for i,j in zip(key, nonkey)]  # finding difference and absolute difference
    diff.sort(key= lambda x: x[0])  # sorting on basis of absolute observed difference
    new_filterlist = list(filter(lambda x: x[0] != 0.0, diff))  # Filtering 0 cases
    rankinglist = list(f_rank(new_filterlist))  # calculating fractional ranks
    newlist = [(r[0]*r[1][1]/r[1][0]) for r in rankinglist]  # calculating Signed ranks
    W_plus = sum(i for i in newlist if i > 0)  # Calculating W+
    W_minus = sum(abs(i) for i in newlist if i < 0)  # Calculating W-

    return min(W_plus, W_minus)  # return minimum of W+ and W-


def open_sheet():
    loc = r'\Algo_Test.xlsx'
    loc = os.getcwd() + loc
    wb = xlrd.open_workbook(loc)  # opening workbook
    return wb.sheet_by_index(0)



if __name__ == '__main__':
    sheet = open_sheet()  # loading first sheet

    key = [sheet.cell_value(r, 1) for r in range(1, sheet.nrows)]  # taking key for 48806

    for i in range(2,sheet.ncols):
        nonkey = [sheet.cell_value(r,i) for r in range(1, sheet.nrows)]  # finding non key one at a time
        print('Value of W for ' + str(int(sheet.cell_value(0, i))) + ' is', str(int(cal_wilcoxon(key, nonkey))))
