#! /usr/bin/env/python

import argparse
import os.path as osp
import shutil

from openpyxl import load_workbook
import pandas as pd
import win32com.client as win32
from xlrd import open_workbook
from xlutils.copy import copy


PARSER = argparse.ArgumentParser(
    description='Update the auto ownership workbooks and uec.')
PARSER.add_argument(
    'iteration', metavar='Iteration', type=int, default=0,
    help='The current calibration iteration number.'
)
PARSER.add_argument(
    '-i', '--input_path', metavar='Input_Path', type=str,
    default='../../newpopsyn',
    help='The path to the directory containing the abm output and uec ' +
    'directories.')
PARSER.add_argument(
    '-o', '--output_path', metavar='Output_Path', type=str, default='.',
    help='The path to the directory containing the model calibration files.')
ARGS = PARSER.parse_args()


def update_auto_ownership(iter_, input_path, output_path):
    """Aggregate model results, calculate constants, and update the UEC.

    Parameters
    ----------
    arg : int
    The calibration iteration number. This number will be appended to the end
    of the aoResults file and 1_AO Calibration file stored in the directory.

    """
    res_path = input_path + '/output/aoResults.csv'
    out = output_path + f'/aoResults-{iter_}.csv'
    shutil.copy2(res_path, out)
    model_results = pd.read_csv(out)
    new_values = model_results.groupby('AO').HHID.count().values
    if iter_ <= 1:
        wb_name = '1_AO Calibration.xlsx'
    else:
        wb_name = f'1_AO Calibration_{iter_ - 1}.xlsx'
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

    workbook.save(f'1_AO Calibration_{iter_}.xlsx')

    excel = win32.gencache.EnsureDispatch('Excel.Application')
    workbook = excel.Workbooks.Open(
        'T:/projects/sr13/develop/2014Calibration/Model Calibration/1_AO/' +
        f'1_AO Calibration_{iter_}.xlsx')
    workbook.Save()
    workbook.Close()
    excel.Quit()

    workbook = load_workbook(f'1_AO Calibration_{iter_}.xlsx', data_only=True)
    auto_ownership = workbook['AO']
    new_constants = [cell.value for cell in auto_ownership['L'][3:8]]
    uec_path = input_path + '/uec/AutoOwnership.xls'
    ao_uec = open_workbook(uec_path, formatting_info=True)
    workbook = copy(ao_uec)
    auto_ownership = workbook.get_sheet(1)
    for idx, val in enumerate(new_constants):
        auto_ownership.write(81, 6 + idx, val)
    workbook.save(uec_path)
    ao_uec.release_resources()


if __name__ == '__main__':
    update_auto_ownership(
        ARGS.iteration, ARGS.input_path, ARGS.output_path)
