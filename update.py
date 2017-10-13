#! /usr/bin/env/python

import argparse
from os.path import abspath
import shutil

from openpyxl import load_workbook
import pandas as pd
# import win32com.client as win32
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


def update(method=None):
    pass

if __name__ == '__main__':
    update()
