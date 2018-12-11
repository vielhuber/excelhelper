<?php
use vielhuber\excelhelper\excelhelper;

class Test extends \PHPUnit\Framework\TestCase
{
    function test__excelhelper()
    {
        $data = [
            ['a1', 'b1', 'c1'],
            ['a2', 'b2', 'c2'],
            [
                ['value' => 'a3', 'background-color' => '#ff0000', 'color' => '#ffffff', 'font-weight' => 'bold', 'border' => '1px solid #000', 'text-align' => 'center'],
                ['value' => 'b3', 'background-color' => '#ff0000', 'color' => '#ffffff', 'font-weight' => 'bold', 'border' => '1px solid #000', 'text-align' => 'left'],
                ['value' => 'c3', 'background-color' => '#ff0000', 'color' => '#ffffff', 'font-weight' => 'bold', 'border' => '1px solid #000', 'text-align' => 'right']
            ]
        ];
        $data = [
            [
                ['value' => 'Partner-ID', 'font-weight' => 'bold'],
                ['value' => 'Firma', 'font-weight' => 'bold'],
                ['value' => 'Anrede', 'font-weight' => 'bold'],
                ['value' => 'Name', 'font-weight' => 'bold'],
                ['value' => 'Vorname', 'font-weight' => 'bold'],
                ['value' => 'RE Straße', 'font-weight' => 'bold'],
                ['value' => 'RE HsNr', 'font-weight' => 'bold'],
                ['value' => 'RE PLZ', 'font-weight' => 'bold'],
                ['value' => 'RE Ort', 'font-weight' => 'bold'],
                ['value' => 'Liefer Straße', 'font-weight' => 'bold'],
                ['value' => 'Liefer HsNr', 'font-weight' => 'bold'],
                ['value' => 'Liefer PLZ', 'font-weight' => 'bold'],
                ['value' => 'Liefer Ort', 'font-weight' => 'bold'],
                ['value' => 'Anbieter', 'font-weight' => 'bold'],
                ['value' => 'Kundennummer', 'font-weight' => 'bold'],
                ['value' => 'Strom/Gas', 'font-weight' => 'bold'],
                ['value' => 'Jahresverbrauch in kWh', 'font-weight' => 'bold'],
                ['value' => 'Lieferbeginn', 'font-weight' => 'bold'],
                ['value' => 'Stichtag', 'font-weight' => 'bold', 'background-color' => '#ffff00'],
                ['value' => 'Zählernummer', 'font-weight' => 'bold'],
                ['value' => 'Hat Zählerwechselwechsel stattgefunden? Neue Zählernummer hier bitte erfassen', 'font-weight' => 'bold', 'background-color' => '#ff0000'],
                ['value' => 'Zählerstand Strom (HT) zum Stichtag OHNE Kommastellen!', 'font-weight' => 'bold', 'background-color' => '#92d050'],
                ['value' => 'Zählerstand Strom (NT) falls vorhanden zum Stichtag OHNE Kommastellen!', 'font-weight' => 'bold', 'background-color' => '#92d050'],
                ['value' => 'Zählerstand Gas zum Stichtag OHNE Kommastellen!', 'font-weight' => 'bold', 'background-color' => '#996633']
            ]
        ];
        excelhelper::write([
            'file' => 'tests/test.xlsx',
            'engine' => 'phpspreadsheet',
            'output' => 'save',
            'data' => $data
        ]);
        $array = excelhelper::read([
            'file' => 'tests/test.xlsx',
            'first_line' => true,
            'format_cells' => false,
            'all_sheets' => false
        ]);
        $this->assertEquals($array, [['a1', 'b1', 'c1'], ['a2', 'b2', 'c2'], ['a3', 'b3', 'c3']]);
    }
}
