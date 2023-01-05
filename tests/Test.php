<?php
use vielhuber\excelhelper\excelhelper;

class Test extends \PHPUnit\Framework\TestCase
{
    function test__excelhelper()
    {
        $data = [
            ['a1', 'b1', 'c1'],
            ['a2', '200€', '1049090014867191'],
            [
                [
                    'value' => 'ca. 200,00 €',
                    'background-color' => '#ff0000',
                    'color' => '#ffffff',
                    'font-weight' => 'bold',
                    'border' => '1px solid #000',
                    'text-align' => 'center'
                ],
                [
                    'value' => '200,50€',
                    'background-color' => '#ff0000',
                    'color' => '#ffffff',
                    'font-weight' => 'bold',
                    'border' => '1px solid #000',
                    'text-align' => 'left'
                ],
                [
                    'value' => '200.50€',
                    'background-color' => '#ff0000',
                    'color' => '#ffffff',
                    'font-weight' => 'bold',
                    'border' => '1px solid #000',
                    'text-align' => 'right'
                ]
            ]
        ];
        excelhelper::write([
            'file' => 'tests/test.xlsx',
            'engine' => 'phpspreadsheet',
            'output' => 'save',
            'style_header' => true,
            'autosize_columns' => false,
            'data' => $data
        ]);
        $array = excelhelper::read([
            'file' => 'tests/test.xlsx',
            'first_line' => true,
            'format_cells' => false,
            'all_sheets' => false,
            'friendly_keys' => false
        ]);
        $this->assertEquals($array, [
            ['a1', 'b1', 'c1'],
            ['a2', '200', '1049090014867191'],
            ['ca. 200,00 €', '200.50', '200.50']
        ]);
        $array = excelhelper::read([
            'file' => 'tests/test.xlsx',
            'first_line' => true,
            'format_cells' => false,
            'all_sheets' => false,
            'friendly_keys' => true
        ]);
        $this->assertEquals($array, [
            1 => ['A' => 'a1', 'B' => 'b1', 'C' => 'c1'],
            2 => ['A' => 'a2', 'B' => '200', 'C' => '1049090014867191'],
            3 => ['A' => 'ca. 200,00 €', 'B' => '200.50', 'C' => '200.50']
        ]);

        excelhelper::write([
            'file' => 'tests/test.xlsx',
            'engine' => 'phpspreadsheet',
            'output' => 'save',
            'style_header' => true,
            'autosize_columns' => false,
            'remove_empty_cols' => true,
            'data' => [
                [3 => 'foo', 7 => 'bar', 12 => '', 14 => '...'],
                [4 => 'bar', 9 => '=2+2', 11 => '', 13 => '...']
            ]
        ]);
        $array = excelhelper::read([
            'file' => 'tests/test.xlsx',
            'first_line' => true,
            'format_cells' => false,
            'all_sheets' => false,
            'friendly_keys' => true
        ]);
        $this->assertEquals($array, [
            1 => ['A' => 'foo', 'B' => 'bar', 'C' => '...'],
            2 => ['A' => 'bar', 'B' => 4, 'C' => '...']
        ]);

        $array = excelhelper::read([
            'file' => 'tests/multi.xlsx',
            'first_line' => true,
            'format_cells' => false,
            'all_sheets' => false,
            'friendly_keys' => false
        ]);
        $this->assertEquals($array, [['a', 'b', 'c']]);
        $array = excelhelper::read([
            'file' => 'tests/multi.xlsx',
            'first_line' => true,
            'format_cells' => false,
            'all_sheets' => true,
            'friendly_keys' => false
        ]);
        $this->assertEquals($array, [
            'foo' => [['a', 'b', 'c']],
            'bar' => [['d', 'e', 'f']],
            'baz' => [['g', 'h', 'i']]
        ]);
    }
}
