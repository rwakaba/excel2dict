import datetime

import excel2dict

def test_reading_book1():
    book1 = excel2dict.to_dict('tests/books/1/Book1.xlsx')
    sheet1 = book1['Sheet1']

    assert len(sheet1) == 10
    assert sheet1[0]['a'] == 1.0
    assert sheet1[0]['b'] == 'qqq'
    assert sheet1[1]['a'] == 2.0
    assert sheet1[1]['b'] == 'aaa'


def test_reading_book2():
    book1 = excel2dict.to_dict('tests/books/2/Book2.xlsx')
    sheet1 = book1['Sheet1']

    assert len(sheet1) == 3

    # int
    assert sheet1[0]['a'] == 1

    # str
    assert sheet1[0]['b'] == 'アイウエオ'

    # bool
    assert sheet1[0]['c'] == True
    assert sheet1[2]['c'] == False

    # date
    assert sheet1[0]['d'] == datetime.date(2019, 7, 12)

    # datetime
    assert sheet1[0]['e'] == datetime.datetime(1900, 2, 27, 23, 59, 59)
    assert sheet1[1]['e'] == datetime.datetime(1900, 2, 28, 0, 0, 0)
    assert sheet1[2]['e'] == datetime.datetime(1900, 2, 28, 23, 59, 59)

    # ref
    assert sheet1[0]['f'][0] == {
        'b': 11,
        'c': 'www',
        'd': {
            'b': 111,
            'c': 'eee'
        }
    }
    assert sheet1[0]['g'] == {
        'b': 11,
        'c': 'www',
        'd': {
            'b': 111,
            'c': 'eee'
        }
    }