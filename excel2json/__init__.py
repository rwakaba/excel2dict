import json

import xlrd

from .helper import *

def expand(sheet_name:str, parsed_all_sheets: dict, dest: dict, jsonify: bool = False):
    sheet_dest = []
    for entry in parsed_all_sheets[sheet_name]:
        entry_dest = {}
        for key, value in entry.items():
            entry_dest[key] = value.expand(parsed_all_sheets, jsonify)
        sheet_dest.append(entry_dest)
    dest[sheet_name] = sheet_dest


def parse_sheet(sheet, configuration=None):
    header = sheet.row_values(0)
    arr = []
    for row_num in range(sheet.nrows - 1):
        row_values = sheet.row_values(row_num + 1)
        d = {}
        for col_num in range(len(header)):
            col_name = header[col_num]
            v = row_values[col_num]
            schema = get_schema(col_name, configuration)
            d[col_name] = Value(v, schema)
        arr.append(d);
    return arr


def to_dict(file_path, configuration: dict = None, jsonify: bool = False):
    wb = xlrd.open_workbook(file_path)
    parsed_all_sheets = {}
    for sheet in wb.sheets():
        setting = None
        if configuration:
            for s in configuration['sheets']:
                if s['name'] == sheet.name:
                    setting = s

        parsed_all_sheets[sheet.name] = parse_sheet(sheet, setting)
    dest = {}
    for sheet_name in parsed_all_sheets.keys():
        expand(sheet_name, parsed_all_sheets, dest, jsonify)
    return dest


def to_json(file_path, configuration: dict = None):
    return to_dict(file_path, configuration, True)


if __name__ == '__main__':
    ret = to_json('Book1.xlsx', load_configuration('excel2json.yaml'));
    print(json.dumps(ret, indent=2, ensure_ascii=False))

