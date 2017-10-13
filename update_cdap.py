#! /usr/bin/env/python

import argparse


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
    pass


if __name__ == '__main__':
    update_cdap(
        ARGS.iteration, ARGS.input_path, ARGS.output_path)
