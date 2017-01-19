#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import xlrd
import csv
import os


def excel2csv(filename):
    """
    Convert an excel file (.xls or .xlsx) to a csv file.

    :param filename: file name (including path) e.g. 'downloads\my_xls.xls'
    """
    wb = xlrd.open_workbook(filename)
    sh = wb.sheet_by_index(0)

    output_file = os.path.splitext(filename)[0] + '.csv'
    your_csv_file = open(output_file, 'w', encoding='utf-8', newline='')
    wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)

    for rownum in range(sh.nrows):
        wr.writerow(sh.row_values(rownum))

    your_csv_file.close()


def batch_convert_xls2csv(path):
    """
    Convert all excel files (.xls and .xlsx) in a specific folder to csv files.

    :param path:  path in which the excel files are located
    """
    for root, dirs, files in os.walk(path):
        for f in files:
            if not f.startswith('.') and f.lower().endswith('.xls') or f.lower().endswith('.xlsx'):
                excel2csv(os.path.join(root,f))
