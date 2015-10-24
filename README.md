# Apache POI Excel parser plugin for Embulk

Parses Microsoft Excel files(xls, xlsx) read by other file input plugins.  
This plugin uses Apache POI.

## Overview

* **Plugin type**: parser
* **Guess supported**: no

## Configuration

* **sheets**: sheet name. (list of string, required)
* **skip_header_lines**: skip rows. (integer, default: `0`)
* **columns**: column definition. see below. (hash, required)

### columns

* **name**: Embulk column name. (string, required)
* **type**: Embulk column type. (string, required)
* **value**: value type. see below. (string, defualt: `cell_value`)
* **column_number**: Excel column number. see below. (string, default: next column)
* **attribute_name**: use with value `cell_style`, `cell_font`, etc. see below. (list of string)
* **on_cell_error**: processing method of Cell error. see below. (string, default: `constant`)
* **on_convert_error**: processing method of convert error. see below. (string, default: `exception`)

### value

* `cell_value`: value in cell.
* `cell_formula`: formula in cell. (if cell is not formula, same `cell_value`.)
* `cell_style`: all cell style attributes. returned json string. see **attribute_name**. (**type** required `string`)
* `cell_font`: all cell font attributes. returned json string. see **attribute_name**. (**type** required `string`)
* `cell_comment`: all cell comment attributes. returned json string. see **attribute_name**. (**type** required `string`)
* `sheet_name`: sheet name.
* `row_number`: row number(1 origin).
* `column_number`: column number(1 origin).
* `constant`: constant value.

  * `constant.`*value*: specified value.
  * `constant`: null.

### column_number

* `A`,`B`,`C`,...: column number of "A1 format".
* *number*: column number (1 origin).
* `+`: next column.
* `+`*name*: next column of name.
* `+`*number*: number next column.
* `-`: previous column.
* `-`*name*: previous column of name.
* `-`*number*: number previous column.
* `=`: same column.
* `=`*name*: same column of name.

### attribute_name

**value**が`cell_style`, `cell_font`, `cell_comment`のとき、デフォルトでは、全属性を取得してJSON文字列に変換します。  
（JSON文字列を返すので、**type**は`string`である必要があります）

```yaml
    columns:
    - {name: foo, type: string, column_number: A, value: cell_style}
```


attribute_nameを指定することで、指定された属性だけを取得してJSON文字列に変換します。

* **attribute_name**: attribute names. (list of string)

```yaml
    columns:
    - {name: foo, type: string, column_number: A, value: cell_style, attribute_name: [border_top, border_bottom, border_left, border_right]}
```


また、`cell_style`や`cell_font`の直後にピリオドを付けて属性名を指定することにより、その属性だけを取得することが出来ます。  
この場合はJSON文字列にはならず、属性の型に合う**type**を指定する必要があります。

```yaml
    columns:
    - {name: foo, type: long, value: cell_style.border}
    - {name: bar, type: long, value: cell_font.color}
```

なお、`cell_style`や`cell_font`では、**column_number**を省略した場合は直前と同じ列を対象とします。  
（`cell_value`では、**column_number**を省略すると次の列に移る）


### on_cell_error

Processing method of Cell error (`#DIV/0!`, `#REF!`, etc).

```yaml
    columns:
    - {name: foo, type: string, column_number: A, value: cell_value, on_cell_error: error_code}
```

* `constant`: set null. (default)
* `constant.`*value*: set value.
* `error_code`: set error code.
* `exception`: throw exception.


### on_convert_error

Processing method of convert error. ex) Excel boolean to Embulk timestamp

```yaml
    columns:
    - {name: foo, type: timestamp, format: "%Y/%m/%d", column_number: A, value: cell_value, on_convert_error: constant.9999/12/31}
```

* `constant`: set null.
* `constant.`*value*: set value.
* `exception`: throw exception. (default)


## Example

```yaml
in:
  type: any file input plugin type
  parser:
    type: poi_excel
    sheets: ["DQ10-orb"]
    skip_header_lines: 1	# first row is header.
    columns:
    - {name: row, type: long, value: row_number}
    - {name: get_date, type: timestamp, value: cell_value, column_number: A}
    - {name: orb_type, type: string}
    - {name: orb_name, type: string}
    - {name: orb_shape, type: long}
    - {name: drop_monster_name, type: string}
```

if omit `value`, specified `cell_value`.  
if omit `column_number` when valus is `cell_value`, specified next column.  
if omit `column_number` when valus is `cell_style`, specified same column.


## Install

```
$ embulk gem install embulk-parser-poi_excel
```


## Build

```
$ ./gradlew package
```
