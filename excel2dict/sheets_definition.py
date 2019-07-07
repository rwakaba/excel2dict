import json
import yaml


class Schema:
    def __init__(self, schema: dict = {}):
        self._def = schema

    @property
    def data_type(self):
        return self._def['type'] if 'type' in self._def else None

    @property
    def ref_sheet(self):
        return self._def['sheet']

    @property
    def source(self):
        return dict(self._def)

    @property
    def is_ref(self):
        return self.data_type == 'ref'

    @property
    def is_array(self):
        return 'is_array' in self._def and self._def['is_array']

    def __repr__(self):
        return f'<Schema source={self._def}/>'

class ColDefinition:
    def __init__(self, definition):
        self._def = definition
        self._schema = Schema(self._def['schema'] if 'schema' in self._def else {})

    @property
    def name(self):
        return self._def['name']

    @property
    def schema(self) -> Schema:
        return self._schema
        # return self._def['schema'] if 'schema' in self._def else None
    
    def __repr__(self):
        return f'<ColDefinition name={self.name}, schema={self.schema}/>'


class SheetDefinition:
    def __init__(self, definition):
        self._def = definition

    @property
    def name(self):
        return self._def['name']

    def get_col_definition(self, col_name) -> ColDefinition:
        if 'columns' in self._def:
            for col_def in self._def['columns']:
                if col_def['name'] == col_name:
                    return ColDefinition(col_def)
                elif 'label' in col_def and col_def['label'] == col_name:
                    return ColDefinition(col_def)
        return ColDefinition({'name': col_name})


def load_sheet_definition(file_path: str) -> dict:
    with open(file_path) as f:
        return yaml.load(f.read(), Loader=yaml.FullLoader)


def get_defenition(sheets_definition, sheet_name: str) -> SheetDefinition:
    for d in sheets_definition['sheets']:
        if d['name'] == sheet_name:
            return SheetDefinition(d)
        elif 'label' in d and d['label'] == sheet_name:
            return SheetDefinition(d)
    return None
