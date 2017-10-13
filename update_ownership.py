#! /usr/bin/env/python

import shutil
import sys

from openpyxl import load_workbook
import pandas as pd
import win32com.client as win32
from xlrd import open_workbook
from xlutils.copy import copy


def update_auto_ownership_uec(arg):
    """Aggregate model results, calculate constants, and update the UEC.

    Parameters
    ----------
    arg : int
    The calibration iteration number. This number will be appended to the end
    of the aoResults file and 1_AO Calibration file stored in the directory.

    """
    shutil.copy2('../../newpopsyn/output/aoResults.csv',
                 'aoResults.csv')
    shutil.copy2('../../newpopsyn/output/aoResults.csv',
                 'aoResults_{}.csv'.format(arg))
    model_results = pd.read_csv('aoResults.csv')
    new_values = model_results.groupby('AO').HHID.count().values
    if arg == '1':
        wb_name = '1_AO Calibration.xlsx'
    else:
        wb_name = '1_AO Calibration_{}.xlsx'.format(int(arg) -1)
    workbook = load_workbook(wb_name, data_only=True)
    ao = workbook['AO']
    prev_constants = []
    for cell in ao['L'][3:8]:
        prev_constants.append(cell.value)
    workbook.close()
    workbook = load_workbook(wb_name)
    ao = workbook['AO']
    for cell, value in zip(ao['K'][3:8], prev_constants):
        cell.value = value
    data = workbook['_data']
    for cell, value in zip(data['B'][1:6], new_values):
        cell.value = value

    workbook.save('1_AO Calibration_{}.xlsx'.format(arg))

    excel = win32.gencache.EnsureDispatch('Excel.Application')
    workbook = excel.Workbooks.Open(
        'T:/projects/sr13/develop/2014Calibration/Model Calibration/1_AO/' +
        '1_AO Calibration_{}.xlsx'.format(arg))
    workbook.Save()
    workbook.Close()
    excel.Quit()

    workbook = load_workbook('1_AO Calibration_{}.xlsx'.format(arg),
                             data_only=True)
    auto_ownership = workbook['AO']
    new_constants = [cell.value for cell in auto_ownership['L'][3:8]]

    ao_uec = open_workbook('../../newpopsyn/uec/AutoOwnership.xls',
                           formatting_info=True)
    workbook = copy(ao_uec)
    auto_ownership = workbook.get_sheet(1)
    for idx, val in enumerate(new_constants):
        auto_ownership.write(81, 6 + idx, val)
    workbook.save('../../newpopsyn/uec/AutoOwnership.xls')
    ao_uec.release_resources()


if __name__ == '__main__':
    update_auto_ownership_uec(sys.argv[1])
