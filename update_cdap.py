#! /usr/bin/env/python

import argparse
import shutil

import pandas as pd


PARSER = argparse.ArgumentParser(
    description='Update the coordinated daily activity pattern calibration '
    'workbooks and uec.')
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


def update_cdap(iter_, input_path, output_path):
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

    """
    uec_path = input_path + '/uec/CoordinatedDailyActivityPattern.xls'
    shutil.copy2(input_path + '/output/personData_3.csv',
                 output_path + f'/personData_{iter_}.csv')
    model_results = pd.read_csv(output_path + f'/aoResults-{iter_}.csv')
    cal_out = output_path + f'/2_CDAP Calibration_{iter_}.xlsx'



if __name__ == '__main__':
    update_cdap(
        ARGS.iteration, ARGS.input_path, ARGS.output_path)
