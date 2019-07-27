# excel2dict

excel2dict support easy loading data from excel files.

## Intalling
```
pip install excel2dict
```

## Quick Over View

Assuming below sample data was saved as excel file named Book.xslx.

|foo|bar|
|--|--|
|'a'|1|
|'b'|2|


Simply, To convert to JSON format text file with command line.

```
$ excel2dict Book.xlsx > out.json
$ less out.json
[
  {
    "foo": 'a',
    "bar": 1 
  },
  {
    "foo": 'b',
    "bar": 2 
  }
]
```

As well, you can do the same thing in python script.

```
>>> import excel2dict
>>> excel2dict.to_dict('Book.xlsx')
[
  {
    "foo": 'a',
    "bar": 1 
  },
  {
    "foo": 'b',
    "bar": 2 
  }
]
```

## Using Sheet Definition
As above example, at simple usage, some data representing dedicated data type(boolean, date, etc) in excel can not be handled usefully.

For this use case, you use a sheet definition file.
if exists, excel2dict load a definition file named `sheet_definition.yaml` from the directory in which target excel file is saved. 

### sample definition
```
sheets:
- name: Members
  cols: 
    - name: member_no
      schema:
        type: int
    - name: name
      schema:
        type: string
    - name: is_active
      schema:
        type: bool
```

### Label Definition
Normally, sheet name is named in a business context in which the name may include multibyte character, space, etc. but for handling in script or JSON text file, named only ascii character is useful.
For this, you can add `label` definition to the definition.

#### For Sheet
```
sheets:
- name: members
  label: New Members
```

#### For Column
```
cols: 
  - name: name
    label: Member's Name
```

### Data Type Definition
excel2dict suppot below data type.

#### int
```
schema:
  type: int
```
#### str
```
schema:
  type: str
```
#### bool
```
schema:
  type: bool
```
#### date
```
schema:
  type: date
```
#### datetime
```
schema:
  type: datetime
```

For needing to adjust timezone, specifing offset is avalable.
```
schema:
  type: datetime
  offset: 9
```
For example, `2019-07-26T09:00:00` in JST, this setting convert the datetime to `2019-07-26T00:00:00`

## A Bit Odd Function
For rare use case, you may need to convert values defined in other sheets as nested structure.

For example, assuming there were 2 sheets as below,  

#### SheetB
|Access Right|Read|Write|
|--|--|--|
|Admin|O|O|
|General|O|X|

#### SheetA
|User|Access Right|
|--|--|
|Scott|Admin|
|Tom|General|

You can get an output like below format defining as `ref` type.
```
[
  {
    "user": 'Scott',
    "group": {
      'read': 'O',
      'write': 'O'
    } 
  },
  {
    "user": 'Tom',
    "group": {
      'read': 'O',
      'write': 'X'
    } 
  }
]
```

## How to specify
Required setting are type and sheet.
- type: `ref`
- sheet: reference sheet name

```
schema:
  type: ref
  sheet: Sheet2
```

For array, specifing `is_array` is avalable.
```
schema:
  type: ref
  sheet: Sheet2
  is_array: true        
```