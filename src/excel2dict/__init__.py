import json

import xlrd

from . import converter, helper

VERSION = (0, 0, 1)

__version__ = '.'.join([str(x) for x in VERSION])

def to_dict(file_path, configuration: dict = None):
    if not configuration:
        configuration = helper.load_configuration('excel2dict.yaml')
    return converter.to_dict(file_path, configuration, False)


def to_json(file_path, configuration: dict = None):
    if not configuration:
        configuration = helper.load_configuration('excel2dict.yaml')
    return converter.to_dict(file_path, configuration, True)

