<?php
use vielhuber\excelhelper\excelhelper;

class Test extends \PHPUnit\Framework\TestCase
{
    function test__excelhelper()
    {
        $data = [
            ['a1', 'b1', 'c1'],
            ['a2', '200€', '374862376723'],
            [
                ['value' => 'a3', 'background-color' => '#ff0000', 'color' => '#ffffff', 'font-weight' => 'bold', 'border' => '1px solid #000', 'text-align' => 'center'],
                ['value' => 'b3', 'background-color' => '#ff0000', 'color' => '#ffffff', 'font-weight' => 'bold', 'border' => '1px solid #000', 'text-align' => 'left'],
                ['value' => 'c3', 'background-color' => '#ff0000', 'color' => '#ffffff', 'font-weight' => 'bold', 'border' => '1px solid #000', 'text-align' => 'right']
            ]
        ];
        excelhelper::write([
            'file' => 'tests/test.xlsx',
            'engine' => 'phpspreadsheet',
            'output' => 'save',
            'style_header' => true,
            'autosize_columns' => true,
            'data' => $data
        ]);
        $array = excelhelper::read([
            'file' => 'tests/test.xlsx',
            'first_line' => true,
            'format_cells' => false,
            'all_sheets' => false
        ]);
        $this->assertEquals($array, [['a1', 'b1', 'c1'], ['a2', '200', '374862376723'], ['a3', 'b3', 'c3']]);
    }
}
