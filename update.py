#! /usr/bin/env/python

import argparse
from os.path import abspath
import shutil

from openpyxl import load_workbook
import pandas as pd
import win32com.client as win32
from xlrd import open_workbook
from xlutils.copy import copy


PARSER = argparse.ArgumentParser(
    description='Update the coordinated daily activity pattern calibration '
    'workbooks and uec.')
PARSER.add_argument(
    'iteration', metavar='Iteration', type=int, default=0,
    help='The current calibration iteration number.'
)
PARSER.add_argument(
    'type', metavar='Type', type=str, default='AO', choices=['AO', 'CDAP'],
    help='The type of update to perform.'
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


def replace_values(dest, data):
    """Replace the values in dest with those in data.

    Parameters
    ----------
    dest : tuple
        Tuple containing openpyxl.cell.cell's.
    data : array
        Array containing data to write into the cells in dest.
    """
    if len(dest) != len(data):
        raise ValueError('Length of dest and data should be the same.')
    for cell, value in zip(dest, data):
        cell.value = value


def update(iter_, input_path, output_path, method='AO'):
    """Aggregate model results, calculate constants, and update the UEC.

    Parameters
    ----------
    iter_ : int
        The calibration iteration number.
    input_path : str
        The relative path to the directory containing the output and uec
        directories.
    output_path : str
        The relative path to the directory containing the calibration files.
    method : str, 'AO' | 'CDAP'
        The type of update to perform.
        Default : 'AO'
        Valid Options :
        - 'AO' : Update AutoOwnership
        - 'CDAP' : Update CoordinatedDailyActivityPattern

    """
    files = {
        'AO': ['AutoOwnership', 'aoResults', '1_AO Calibration',
               update_ao],
        'CDAP': ['CoordinatedDailyActivityPattern', 'personData_3',
                 '2_CDAP Calibration', update_cdap]}

    cal_path = output_path + f'/{files[method][2]}_{iter_}.xlsx'
    uec_path = input_path + f'/uec/{files[method][0]}.xls'
    shutil.copy2(input_path + f'/output/{files[method][1]}.csv',
                 output_path + f'/{files[method][1]}_{iter_}.csv')
    results = pd.read_csv(output_path + f'/{files[method][1]}_{iter_}.csv')
    if iter_ < 1:
        wb_name = output_path + f'/{files[method][2]}.xlsx'
    else:
        wb_name = output_path + f'/{files[method][2]}_{iter_ - 1}.xlsx'
    files[method][3](iter_, wb_name, results, uec_path, cal_path)



if __name__ == '__main__':
    update()
