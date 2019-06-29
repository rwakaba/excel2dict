import json
import yaml

import xlrd

setting_src = """
    sheets:
    - name: Sheet1
      cols: 
        - name: a
          label: COL A
          schema:
            type: int
        - name: b
          label: COL B
          schema:
            type: string
        - name: c
          label: COL C
          schema:
            type: datetime
        - name: d
          label: COL D
          schema:
            type: datetime
      dependencies: 
        - hoge
        - moge
        
"""
settings = yaml.load(setting_src)
print(settings)


def get_schema(name, setting):
    for col in setting['cols']:
        if col['name'] == name:
            return col
    return None


def format(src_value, col_definition: dict):
    schema = col_definition['schema']
    if schema['type'] == 'int':
        return int(src_value)
    return src_value


def parse_sheet(sheet, setting=None):
    header = sheet.row_values(0)
    arr = []
    for row_num in range(sheet.nrows - 1):
        row_values = sheet.row_values(row_num + 1)
        d = {}
        for col_num in range(len(header)):
            name = header[col_num]
            v = row_values[col_num]
            if setting:
                v = format(v, get_schema(name, setting))
                # schema = get_schema(name, setting)['schema']
                # if schema['type'] == 'int':
                #     v = int(v)
            d[name] = v
        arr.append(d);
    return arr


def to_json():
    wb = xlrd.open_workbook('Book1.xlsx')
    arr = []
    for sheet in wb.sheets():
        setting = None
        for s in settings['sheets']:
            if s['name'] == sheet.name:
                setting = s

        arr.append({
            'sheet': sheet.name,
            'data': parse_sheet(sheet, setting),
        })
    return arr


if __name__ == '__main__':
    ret = to_json();
    print(json.dumps(ret, indent=2, ensure_ascii=False))

