import argparse
import datetime

import yaml

from . import serial_value_handler

def parse_args():
    parser = argparse.ArgumentParser("excel2dict")
    parser.add_argument('target_file', type=str, help="target excel file")
    parser.add_argument('--definition', type=str, help="path to sheet definition file")
    return parser.parse_args()


def format(src_value: str, schema: dict, jsonify: bool = False):
    if not schema:
        return src_value
    tp = schema['type']
    if tp == 'int':
        return int(src_value)
    if tp == 'str':
        return str(src_value)
    if tp == 'bool':
        if type(src_value) != int or not (0 <= src_value <=1):
            raise Exception(f'the value `{src_value}`` is not applicable for type {tp}')
        return True if src_value == 1 else False
    if tp == 'date':
        if type(src_value) != float:
            raise Exception(f'the value `{src_value}`` is not applicable for type {tp}')
        dt = serial_value_handler.serial2date(src_value)
        if jsonify:
            # TODO to use format specified in yaml configuration
            return dt.isoformat()
        else:
            return dt
    if tp == 'datetime':
        if type(src_value) != float:
            raise Exception(f'the value `{src_value}`` is not applicable for type {tp}')
        offset = schema['offset'] if 'offset' in schema else 0
        dt = serial_value_handler.serial2datetime(src_value, offset)
        if jsonify:
            # TODO to use format specified in yaml configuration
            return dt.isoformat()
        else:
            return dt

    return src_value

