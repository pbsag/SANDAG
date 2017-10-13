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

    results = pd.read_csv(output_path + f'/aoResults-{iter_}.csv')
    res_vals = results.groupby(['type', 'activity_pattern']).size()\
        .reset_index()
    res_vals.loc[res_vals.type == 'Child too young for school', 'type'] = \
        'Pre-school'
    res_vals.loc[res_vals.type == 'Non-worker', 'type'] = 'Non-working Adult'
    res_vals.loc[res_vals.type == 'Retired', 'type'] = 'Non-working Senior'
    res_vals.loc[res_vals.type == 'Student of driving age', 'type'] = \
        'Driving Age Student'
    res_vals.loc[res_vals.type == 'Student of non-driving age', 'type'] = \
        'Non-driving Student'
    res_vals.loc[res_vals.type == 'University Student', 'type'] = \
        'College Student'
    res_vals = res_vals.sort_values(['type', 'activity_pattern'])[0].values



if __name__ == '__main__':
    update_cdap(
        ARGS.iteration, ARGS.input_path, ARGS.output_path)
