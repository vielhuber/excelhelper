# 📗 excelhelper 📗

excelhelper is a helper for excel.

## Installation

install once with [composer](https://getcomposer.org/):

```
composer require vielhuber/excelhelper
```

then add this to your files:

```php
require __DIR__.'/vendor/autoload.php';
use vielhuber\excelhelper\excelhelper;
```

## Usage

```php
excelhelper::write([
    'engine' => 'phpspreadsheet', // phpspreadsheet|spout
    'file' => 'file.xlsx', // can write xlsx, xls and csv
    'output' => 'save', // save|download
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

excelhelper::read([
    'file' => 'file.xlsx',
    'first_line' => true, // true|false
    'format_cells' => false, // false|true
    'all_sheets' => false, // false|true
]);
```