import json
import xlrd

from excel2dict import helper
from excel2dict.sheets_definition import (
    load_sheet_definition,
    get_defenition,
    SheetDefinition,
    ColDefinition,
    Schema
)


def to_dict(file_path, sheets_definition: dict = None, jsonify: bool = False):
    wb = xlrd.open_workbook(file_path)
    parsed_all_sheets = {}
    for sheet in wb.sheets():
        definition = None
        sheet_name = sheet.name
        if sheets_definition:
            definition: SheetDefinition = get_defenition(sheets_definition, sheet_name)
            if definition:
                sheet_name = definition.name

        parsed_all_sheets[sheet_name] = read_sheet(sheet, definition)
    dest = {}
    for sheet_name in parsed_all_sheets.keys():
        expand(sheet_name, parsed_all_sheets, dest, jsonify)
    return dest


def read_sheet(sheet, sheet_definition: SheetDefinition = None):
    header = sheet.row_values(0)
    arr = []
    for row_num in range(sheet.nrows - 1):
        row_values = sheet.row_values(row_num + 1)
        d = {}
        for col_num in range(len(header)):
            col_name = header[col_num]
            schema = None
            if sheet_definition:
                col_def: ColDefinition = sheet_definition.get_col_definition(col_name)
                col_name = col_def.name
                schema = col_def.schema

            v = row_values[col_num]
            d[col_name] = Value(v, schema)
        arr.append(d);
    return arr


def expand(sheet_name:str, parsed_all_sheets: dict, dest: dict, jsonify: bool = False):
    sheet_dest = []
    for entry in parsed_all_sheets[sheet_name]:
        entry_dest = {}
        for key, value in entry.items():
            """
            a value is `a` -> return `a`
            a value is `ref1` is defined in anoter sheet-> return `{'foo': 1, 'bar': 'sample 1'}`
            a value is `ref1, ref2` is defined in anoter sheet-> return `[{'foo': 1, 'bar': 'sample 1}, {'foo': 2, 'bar': 'sample 2}]`
            """
            entry_dest[key] = value.expand(parsed_all_sheets, jsonify)
        sheet_dest.append(entry_dest)
    dest[sheet_name] = sheet_dest


class Value:
    def __init__(self, raw: any, schema: Schema):
        self.raw = raw
        self.schema = schema
        self.sheet_name = None
        self._value = None
        self._to_ref()

    def _to_ref(self):
        schema = self.schema
        if schema.is_ref:
            self.sheet_name = schema.ref_sheet
            src = self.raw
            if schema.is_array:
                self.ref = map(lambda x: x.strip(), src.split(','))
            else:
                self.ref = src

    def expand(self, parsed_all_sheets: dict, jsonify: bool = False):
        if self._value:
            return self._value
        if self.schema.is_ref:
            self._value = self._expand(parsed_all_sheets, jsonify)
        else:
            self._value = helper.format(self.raw, self.schema.source, jsonify)
        return self._value

    def _expand(self, parsed_all_sheets: dict, jsonify: bool = False) -> any:
        ref_sheet_name = self.sheet_name
        ref_sheet = parsed_all_sheets[ref_sheet_name]

        def _search(ref_name):
            for row in ref_sheet:
                """
                row contents
                {
                    'ref_name': <Value raw=ref1 />,
                    'col_a': <Value raw=xxx />,
                    'col_b': <Value raw=xxx />
                }
                """
                if row['ref_name'].raw == ref_name:
                    return row
            raise Exception(f'no such data `{ref_name}` in {self.sheet_name}')

        def _expand_row(row: dict):
            expanded = {}
            for k, v in row.items():
                value: Value = v
                if k != 'ref_name':
                    expanded[k] = value.expand(parsed_all_sheets, jsonify)
            return expanded

        if self.schema.is_array:
            arr = []
            for ref_name in self.ref:  # ['re1', 'ref2', 'ref3', ..]
                arr.append(_expand_row(_search(ref_name)))
            return arr
        else:
            return _expand_row(_search(self.ref))

    def __repr__(self):
        return f'<Value raw={self.raw}, expanded={self._value} />'

if __name__ == '__main__':
    args = helper.parse_args()
    definition = None
    if args.definition:
        definition = load_sheet_definition(args.definition)
    ret = to_dict(args.target_file, definition, True)
    print(json.dumps(ret['Sheet1'], indent=2, ensure_ascii=False))