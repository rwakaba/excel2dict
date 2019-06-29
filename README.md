# excel2json

```
> excel2json -f Book.xls > out.json
```

### simple sheet
|foo|bar|baz|
|--|--|--|
|'a'|1|TRUE|
|'b'|2|FALSE|


```
>>> import excel2json

>>> excel2json.to_json('Book.xlsx')
[
  {
    'foo': 'a',
    'bar': 1.0,
    'baz': 1,
  },
  {
    'foo': 'b',
    'bar': 2.0,
    'baz': 0,
  }
]
```

### configulation
You can use a configuration file that describe header's label, data type of column, and so.

```
sheets:
- name: Sheet1
  cols: 
    - name: foo
      label: COL A
      schema:
        type: int
    - name: bar
      label: COL B
      schema:
        type: string
    - name: baz
      label: COL C
      schema:
        type: bool
    - name: bazz
      label: COL D
      schema:
        type: datetime
  dependencies: 
    - Sheet2
```
        