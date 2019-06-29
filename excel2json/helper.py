import datetime
from decimal import *

import yaml

def load_configuration(file_path: str) -> dict:
    # path to excel2json.yaml
    with open(file_path) as f:
        return yaml.load(f.read())


def get_col_definition(col_name, configuration: dict = None):
    if not configuration:
        return None
    for col in configuration['columns']:
        if col['name'] == col_name:
            return col
    return None


def get_schema(col_name, configuration: dict = None):
    if not configuration:
        return None
    col_definition = get_col_definition(col_name, configuration)
    return col_definition['schema'] if col_definition else None


def serial2delta(d: int = 0, t: int = 0, offset = 0):
    # delta of 1 means 1900-01-01 and `delta_d` means offset from 1900-01-01
    # therfore `delta_d` needs to minus 1
    delta_d = d - 1
    delta_h = -1 * offset
    delta_t = Decimal(float('0.' + str(t))) * Decimal(24) * Decimal(3600)
    delta_t = round(delta_t)
    if 60 < d:
        # excel has the date 1900 Feb 29th that does not exist. 
        # then `delta_d` needs to minus 1 if d is after 1900-02-28(value of d is 60)
        delta_d = delta_d - 1
    return datetime.timedelta(days=delta_d, hours=delta_h, seconds=delta_t)


def serial2date(src_value: int):
    delta_d = src_value - 2
    delta = datetime.timedelta(days=delta_d)
    return datetime.date(1900, 1, 1) + serial2delta(src_value)


def serial2datetime(src_value: float, offset = 0):
    date_time = str(src_value).split('.')
    d_part = int(date_time[0])
    t_part = int(date_time[1])
    return datetime.datetime(1900, 1, 1) + serial2delta(d_part, t_part, offset)


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
    if tp == 'datetime':
        if type(src_value) != float:
            raise Exception(f'the value `{src_value}`` is not applicable for type {tp}')
        offset = schema['offset'] if 'offset' in schema else 0
        dt = serial2datetime(src_value, offset)
        if jsonify:
            # TODO to use format specified in yaml configuration
            return dt.isoformat()
        else:
            return dt

    return src_value


class Value:
    def __init__(self, raw, schema: dict = None):
        self.raw = raw
        self.schema = schema
        self.sheet_name = None
        self._value = None
        self._to_ref()

    def _to_ref(self):
        src = self.raw
        if type(src) == str and '!' in src:
            sheet_name, refs = src.split('!')
            self.sheet_name = sheet_name
            if '[' in refs:
                self.ref = refs.strip('[').strip(']').split(',')
                self.json_type =  list
            elif '{' in refs:
                self.ref = refs.strip('{').strip('}')
                self.json_type = dict

    def expand(self, parsed_all_sheets: dict, jsonify: bool = False):
        if self._value:
            return self._value
        if self.sheet_name and self.sheet_name in parsed_all_sheets.keys():
            self._value = self._expand(parsed_all_sheets, jsonify)
        else:
            self._value = self._format_value(self.raw, jsonify)
        return self._value

    def _format_value(self, src_value, jsonify: bool = False):
        return format(src_value, self.schema, jsonify)

    def _expand(self, parsed_all_sheets: dict, jsonify: bool = False) -> any:
        """
        [
          {
            "ref_name": Value('ref1'),
            "a": Value(1.0),
            "b": Value("qqq"),
            "c": Value(43645.5),
            "d": Value(43645.541655092595),
            "e": Value("Sheet2![ref1]")
          },
        ],
        """
        ref_sheet_name = self.sheet_name
        ref_sheet = parsed_all_sheets[ref_sheet_name]
        if self.json_type == list:
            arr = []
            for ref_name in self.ref:
                b = False
                for row in ref_sheet:
                    if row['ref_name'].raw == ref_name:
                        b = True
                        reduced = {}
                        for k, v in row.items():
                            if k != 'ref_name':
                                reduced[k] = v.expand(parsed_all_sheets, jsonify)
                        arr.append(reduced)
                        break
                if not b:
                    raise Exception(f'no such name in {ref_sheet_name}')
            return arr
        elif self.json_type == dict:
            ref_name = self.ref
            for row in ref_sheet:
                if row['ref_name'].raw == ref_name:
                    reduced = {}
                    for k, v in row.items():
                        if k != 'ref_name':
                            reduced[k] = v.expand(parsed_all_sheets, jsonify)
                    return reduced
            raise Exception(f'no such name in {ref_sheet_name}')
        else:
            raise Exception()


