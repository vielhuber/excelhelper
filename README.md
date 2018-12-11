# 📗 excelhelper 📗

excelhelper is a helper for excel.

with its help you can write and read to xlsx/xls/csv in php in a very simple way.

## installation

install once with [composer](https://getcomposer.org/):

```
composer require vielhuber/excelhelper
```

then add this to your files:

```php
require __DIR__.'/vendor/autoload.php';
use vielhuber\excelhelper\excelhelper;
```

## usage

### reading

```php
$array = excelhelper::read([
    'file' => 'file.xlsx',
    'first_line' => true, // true|false
    'format_cells' => false, // false|true
    'all_sheets' => false, // false|true
]);
```

### writing

```php
excelhelper::write([
    'file' => 'file.xlsx', // can write xlsx, xls and csv
    'engine' => 'phpspreadsheet', // phpspreadsheet|spout
    'output' => 'save', // save|download
    'style_header' => true, // true|false
    'autosize_columns' => true, // true|false
    'data' => [
        ['a1','b1','c1'],
        ['a2','b2','c2'],
        [
          ['value' => 'a3', 'background-color': '#ff0000', 'color': '#ffffff', 'font-weight': 'bold', 'border': '1px solid #000', 'text-align': 'center'],
          ['value' => 'b3', 'background-color': '#ff0000', 'color': '#ffffff', 'font-weight': 'bold', 'border': '1px solid #000', 'text-align': 'left'],
          ['value' => 'c3', 'background-color': '#ff0000', 'color': '#ffffff', 'font-weight': 'bold', 'border': '1px solid #000', 'text-align': 'right'],
        ]
    ]
]);
```
