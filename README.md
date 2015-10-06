# Apache POI Excel parser plugin for Embulk

Parses Microsoft Excel files(xls, xlsx) read by other file input plugins.  
This plugin uses Apache POI.

## Overview

* **Plugin type**: parser
* **Guess supported**: no

## Configuration

* **sheet**: sheet name. (string, default: `Sheet1`)
* **skip_header_lines**: skip rows. (integer, default: `0`)
* **columns**: column definition. see below. (hash, required)

### columns

* **name**: Embulk column name. (string, required)
* **type**: Embulk column type. (string, required)
* **value_type**: value type. see below. (string, defualt: `cell_value`)
* **column_number**: Excel column number. `A`,`B`,`C`,... or number(1 origin). (string, default: next column)

### value_type

* **cell_value**: value in cell.
* **cell_formula**: formula in cell. (if cell is not formula, same `cell_value`.)
* **sheet_name**: sheet name.
* **row_number**: row number(1 origin).


## Example

```yaml
in:
  type: any file input plugin type
  parser:
    type: poi_excel
    sheet: "DQ10-orb"
    skip_header_lines: 1	# first row is header.
    columns:
    - {name: row, type: long, value_type: row_number}
    - {name: get_date, type: timestamp, value_type: cell_value, column_number: A}
    - {name: orb_type, type: string}
    - {name: orb_name, type: string}
    - {name: orb_shape, type: long}
    - {name: drop_monster_name, type: string}
```

if omit `value_type`, specified `cell_value`.  
if omit `column_number`, specified next column.

### execute

```
$ cd ~/your-workspace
$ git clone https://github.com/hishidama/embulk-parser-poi_excel.git
$ cd embulk-parser-poi_excel
$ ./gradlew package
$ cd /your-embulk-working-dir
$ embulk run -L ~/your-workspace/embulk-parser-poi_excel config.yml
```

## Build

```
$ ./gradlew package
```
