import argparse
import datetime

import yaml


def parse_args():
    parser = argparse.ArgumentParser("excel2dict")
    parser.add_argument('target_file', type=str, help="target excel file")
    parser.add_argument('-s', '--sheet-definition', type=str, help="path to sheet definition file")
    parser.add_argument('-c', '--compact-output', action='store_true', help="compact output instead of pretty JSON format.")
    parser.add_argument('-i', '--indent', type=int, default=2, help="number of spaces")
    parser.add_argument('-a', '--ascii-output', action='store_true', help="ensure ascii code points")
    return parser.parse_args()
