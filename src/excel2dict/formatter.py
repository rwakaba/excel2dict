import datetime

from excel2dict import serial_value_handler


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
            if 'format' in schema:
                return dt.strftime(schema['format'])
            else:
                return dt.isoformat()
        else:
            return dt
    if tp == 'datetime':
        if type(src_value) != float:
            raise Exception(f'the value `{src_value}`` is not applicable for type {tp}')
        offset = schema.get('offset', 0)
        dt = serial_value_handler.serial2datetime(src_value, offset)
        if jsonify:
            if 'format' in schema:
                return dt.strftime(schema['format'])
            else:
                return dt.isoformat()
        else:
            return dt

    return src_value
