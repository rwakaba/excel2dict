import json

import xlrd

from excel2dict import converter
from excel2dict.sheets_definition import load_sheet_definition

VERSION = (0, 0, 1)

__version__ = '.'.join([str(x) for x in VERSION])

def to_dict(target_file_path, sheets_definition: dict = None, jsonify: bool = False):
    if not sheets_definition:
        import os
        f = f'{os.path.dirname(target_file_path)}/sheet_definition.yaml'
        if os.path.exists(f):
            sheets_definition = load_sheet_definition(f)
    return converter.to_dict(target_file_path, sheets_definition, jsonify)


def to_json(file_path, sheets_definition: dict = None):
    return to_dict(file_path, sheets_definition, True)

