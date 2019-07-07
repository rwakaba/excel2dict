import datetime
from decimal import *


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
    return datetime.date(1900, 1, 1) + serial2delta(src_value)


def serial2datetime(src_value: float, offset = 0):
    date_time = str(src_value).split('.')
    d_part = int(date_time[0])
    t_part = int(date_time[1])
    return datetime.datetime(1900, 1, 1) + serial2delta(d_part, t_part, offset)