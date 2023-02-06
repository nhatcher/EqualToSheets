# equalto

Python package allowing to create, edit and evaluate Excel spreadsheets.

## Requirements

Make sure you have [Python 3.9+](https://docs.python.org/3.9/) and [pip](https://pypi.org/project/pip/) installed.

## Installation

Optionally create and activate a python virtual environment:

```bash
pip3 install --upgrade pip virtualenv
python3 -m virtualenv --clear venv
source venv/bin/activate
```

Install the package (it's important that the find-links parameter points to the directory with *.whl files).

```bash
pip install equalto --no-index --find-links equalto
```

Launch a python shell with the equalto package available:

```
$ venv/bin/python
Python 3.9.16 (main, Dec  7 2022, 01:11:58) 
[GCC 7.5.0] on linux
Type "help", "copyright", "credits" or "license" for more information.
>>> import equalto
```

## Quickstart

### Opening an existing .xlsx file

```python
>>> import equalto
>>> wb = equalto.load('filename.xlsx')
```

### Creating a new workbook

```python
>>> import equalto

# creates a new workbook using the local time zone
>>> wb = equalto.new()

# `timezone` parameter can be used to specify the workbook time zone
>>> from zoneinfo import ZoneInfo
>>> wb = equalto.new(timezone=ZoneInfo('US/Central'))
```

### Exporting workbooks to .xlsx files

```python
>>> wb.save('output.xlsx')
```

### Workbook properties

```python
>>> wb = equalto.new(timezone=ZoneInfo('Europe/Berlin'))
>>> wb.timezone
zoneinfo.ZoneInfo(key='Europe/Berlin')
```

### `Sheet` class

#### Loading sheets

Sheets can be accessed by their name:

```python
>>> sheet = wb.sheets['Sheet1']
```

Sheets can be accessed by their index:

```python
>>> first_sheet = wb.sheets[0]
>>> last_sheet = wb.sheets[-1]
```

#### Creating sheets

```python
# creates sheet named 'Calculation'
>>> calc_sheet = wb.sheets.add('Calculation')

# the sheet name is optional, defaults to Sheet1, Sheet2...
>>> sheet = wb.sheets.add()
```

#### Renaming sheets

```python
>>> sheet.name
'Sheet1'
>>> sheet.name = 'Data'
```

#### Deleting sheets

```python
# deletes 'Calculation' sheet
>>> del wb.sheets['Calculation']

# deletes `sheet`
>>> sheet.delete()
```

### `Cell` class

#### Cells can be accessed globally via `Workbook` instance

```python
>>> a1_cell = wb['Sheet1!A1']
>>> a2_cell = wb.cell(sheet_index=0, row=2, column=1)
```

#### Cells can be accessed at the sheet level via `Sheet` instance

```python
>>> sheet = wb.sheets['Sheet1']
>>> a1_cell = sheet['A1']
>>> a2_cell = sheet.cell(row=2, column=1)
```

#### `value` property

```python
# returns unformatted raw value (float | bool | str | None)
>>> wb['Sheet1!A1'].value
42.0
```

There are also convenience properties that perform the type check and necessary conversions:

```python
>>> wb['Sheet1!A1'].int_value
7

>>> wb['Sheet1!A2'].float_value
4.2

>>> wb['Sheet1!A3'].str_value
'Header'

>>> wb['Sheet1!A4'].date_value
datetime.date(2023, 1, 11)

>>> wb['Sheet1!A5'].datetime_value
datetime.datetime(2023, 1, 11, 9, 3, 23, 16575, tzinfo=zoneinfo.ZoneInfo(key='Europe/Berlin'))

>>> wb['Sheet1!A6'].bool_value
True
```

If a type conversion is not possible for a given convenience property, then an exception is raised:

```python
>>> cell.value = 4.2
>>> cell.int_value
equalto.exceptions.WorkbookValueError: 4.2 is not an integer
```

Value setter:

```python
>>> wb['Sheet1!A1'].value = 42
```

Value setter automatically handles type conversions:

```python
>>> from datetime import date
>>> date_cell.value = date(2022, 12, 1)
# `date_cell.value` returns the float value 44896.0, since Excel stores dates as
# the ~number of days since 1900.
>>> date_cell.value
44896.0
```

```python
>>> from datetime import datetime
>>> from zoneinfo import ZoneInfo
>>> timestamp_cell.value = datetime(2023, 1, 10, 10, tzinfo=ZoneInfo('Europe/Berlin'))
# `timestamp_cell.value` returns the float value 44936.375, since Excel stores timestamps as
# the ~number of days since 1900, with time represented as a fraction of a day.
>>> timestamp_cell.value  # numeric representation, `value` is raw excel type
44936.375
```

#### `Cell.__str__` implementation respects the cell formatting

```python
>>> cell = wb.sheets['Sheet1']['A1']
>>> cell.style.format = "$#,##0.00"
>>> cell.value = 12345.67
>>> str(cell)
'$12,345.67'
```

#### `formula` property

```python
>>> wb['Sheet1!A2'].value = 1
>>> wb['Sheet1!A3'].value = 2
>>> cell = wb['Sheet1!A1']
>>> cell.formula  # None
>>> cell.formula = '=A2+A3'
>>> cell.formula
'=A2+A3'
>>> cell.value  # computed value
3.0
```

Setting value explicitly removes the formula:

```python
>>> cell.value = 42
>>> cell.formula
None
```

#### `type` property aligned with TYPE() function in Excel

```python
class CellType(Enum):
    number = 1
    text = 2
    logical_value = 4
    error_value = 16
    array = 64
    compound_data = 128
```

```python
>>> cell.value = 7
>>> cell.type
<CellType.number: 1>

>>> cell.formula = '=1/0'
>>> cell.type
<CellType.error_value: 16>
```

#### `style` property

Reading/setting a cell formatting:

```python
>>> cell.style.format
'general'
>>> cell.value = 41343
>>> str(cell)
'41343'
>>> cell.style.format = '$#,##0.00'
>>> str(cell)
'$41,343.00'
>>> cell.style.format = 'yyyy-mm-dd'
>>> str(cell)
'2013-03-10'
```

#### Deleting a cell

```python
# deletes the value of the cell
>>> wb['Sheet1!A1'].value = None

# deletes the value, format and style
>>> del wb['Sheet1!A1']

# deletes the value, format and style
>>> cell = wb['Sheet1!A1']
>>> cell.delete()
```

