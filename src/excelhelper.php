<?php
namespace vielhuber\excelhelper;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;

class excelhelper
{
    public static function read($args)
    {
        $args = self::prepareReadArgs($args);

        $sheets = [];
        $filetype = IOFactory::identify($args['file']);
        $reader = IOFactory::createReader($filetype);
        $spreadsheet = $reader->load($args['file']);

        if ($args['all_sheets'] === false) {
            $sheet = $spreadsheet->getSheet(0);
            $array = $sheet->toArray(null, true, $args['format_cells']);
            if ($args['first_line'] === false) {
                unset($array[0]);
                $array = array_values($array);
            }
            $sheets[$sheet->getTitle()] = $array;
        } else {
            foreach ($spreadsheet->getWorksheetIterator() as $sheet_nr => $sheet) {
                $array = $sheet->toArray(null, true, $args['format_cells']);
                if ($args['first_line'] === false) {
                    unset($array[0]);
                    $array = array_values($array);
                }
                $sheets[$sheet->getTitle()] = $array;
            }
        }

        if ($args['friendly_keys'] === true) {
            foreach ($sheets as $sheets__key => $array) {
                $array_friendly = [];
                foreach ($array as $array__key => $array__value) {
                    $array_friendly[$array__key + 1] = [];
                    foreach ($array__value as $array__value__key => $array__value__value) {
                        $array_friendly[$array__key + 1][self::int2char($array__value__key + 1)] = $array__value__value;
                    }
                }
                $sheets[$sheets__key] = $array_friendly;
            }
        }

        if (count($sheets) <= 1) {
            $sheets = $sheets[array_key_first($sheets)];
        }

        return $sheets;
    }

    private static function prepareReadArgs($args)
    {
        // default values
        if (!array_key_exists('first_line', $args)) {
            $args['first_line'] = true;
        }
        if (!array_key_exists('format_cells', $args)) {
            $args['format_cells'] = false;
        }
        if (!array_key_exists('convert_bignumbers', $args)) {
            $args['convert_bignumbers'] = true;
        }
        if (!array_key_exists('all_sheets', $args)) {
            $args['all_sheets'] = false;
        }
        if (!array_key_exists('friendly_keys', $args)) {
            $args['friendly_keys'] = true;
        }
        // checks
        if (!isset($args['file']) && !file_exists($args['file'])) {
            throw new \Exception('file missing');
        }
        if (!in_array($args['first_line'], [true, false])) {
            throw new \Exception('unknown first_line');
        }
        if (!in_array($args['format_cells'], [true, false])) {
            throw new \Exception('unknown format_cells');
        }
        if (!in_array($args['convert_bignumbers'], [true, false])) {
            throw new \Exception('unknown convert_bignumbers');
        }
        if (!in_array($args['all_sheets'], [true, false])) {
            throw new \Exception('unknown all_sheets');
        }
        if (!in_array($args['friendly_keys'], [true, false])) {
            throw new \Exception('unknown friendly_keys');
        }
        return $args;
    }

    public static function write($args)
    {
        $args = self::prepareWriteArgs($args);
        if ($args['engine'] === 'phpspreadsheet') {
            $spreadsheet = new Spreadsheet();
            $sheet = $spreadsheet->getActiveSheet();

            foreach ($args['data'] as $data__key => $data__value) {
                $row = $data__key + 1;
                foreach ($data__value as $data__value__key => $data__value__value) {
                    $col = self::int2char($data__value__key + 1);
                    if (!is_array($data__value__value)) {
                        /* big numbers */
                        if (is_numeric($data__value__value) && strlen($data__value__value) > 10) {
                            $data__value__value = '\'' . $data__value__value;
                        }
                        $sheet->setCellValue($col . $row, $data__value__value);
                        $sheet
                            ->getStyle($col . $row)
                            ->getAlignment()
                            ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT);
                    } else {
                        if (array_key_exists('value', $data__value__value)) {
                            /* big numbers */
                            if (is_numeric($data__value__value['value']) && strlen($data__value__value['value']) > 10) {
                                $data__value__value['value'] = '\'' . $data__value__value['value'];
                            }
                            $sheet->setCellValue($col . $row, $data__value__value['value']);
                        }
                        if (array_key_exists('background-color', $data__value__value)) {
                            $sheet
                                ->getStyle($col . $row)
                                ->getFill()
                                ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
                            $sheet
                                ->getStyle($col . $row)
                                ->getFill()
                                ->getStartColor()
                                ->setARGB('00' . str_replace('#', '', $data__value__value['background-color']));
                        }
                        if (array_key_exists('color', $data__value__value)) {
                            $sheet
                                ->getStyle($col . $row)
                                ->getFont()
                                ->getColor()
                                ->setARGB('00' . str_replace('#', '', $data__value__value['color']));
                        }
                        if (
                            array_key_exists('font-weight', $data__value__value) &&
                            $data__value__value['font-weight'] === 'bold'
                        ) {
                            $sheet
                                ->getStyle($col . $row)
                                ->getFont()
                                ->setBold(true);
                        }
                        if (
                            array_key_exists('text-align', $data__value__value) &&
                            $data__value__value['text-align'] === 'center'
                        ) {
                            $sheet
                                ->getStyle($col . $row)
                                ->getAlignment()
                                ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
                        }
                        if (
                            array_key_exists('text-align', $data__value__value) &&
                            $data__value__value['text-align'] === 'right'
                        ) {
                            $sheet
                                ->getStyle($col . $row)
                                ->getAlignment()
                                ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT);
                        }
                        if (
                            !array_key_exists('text-align', $data__value__value) ||
                            $data__value__value['text-align'] === 'left'
                        ) {
                            $sheet
                                ->getStyle($col . $row)
                                ->getAlignment()
                                ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT);
                        }
                        if (array_key_exists('border', $data__value__value)) {
                            $sheet
                                ->getStyle($col . $row)
                                ->getBorders()
                                ->getTop()
                                ->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
                            $sheet
                                ->getStyle($col . $row)
                                ->getBorders()
                                ->getLeft()
                                ->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
                            $sheet
                                ->getStyle($col . $row)
                                ->getBorders()
                                ->getBottom()
                                ->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
                            $sheet
                                ->getStyle($col . $row)
                                ->getBorders()
                                ->getRight()
                                ->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
                        }
                    }
                }
            }

            if ($args['style_header'] === true) {
                $sheet
                    ->getStyle('A1:' . $sheet->getHighestColumn() . '1')
                    ->getFont()
                    ->setBold(true);
                $sheet
                    ->getStyle('A1:' . $sheet->getHighestColumn() . '1')
                    ->getAlignment()
                    ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
                $sheet
                    ->getStyle('A1:' . $sheet->getHighestColumn() . '1')
                    ->getAlignment()
                    ->setVertical(\PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER);
                $sheet
                    ->getStyle('A1:' . $sheet->getHighestColumn() . '1')
                    ->getAlignment()
                    ->setWrapText(true);
                $sheet->getRowDimension('1')->setRowHeight(60);
            }

            if ($args['autosize_columns'] === true) {
                // first set columns to auto size
                $toCol = $sheet->getColumnDimension($sheet->getHighestColumn())->getColumnIndex();
                $toCol++;
                for ($i = 'A'; $i !== $toCol; $i++) {
                    $sheet->getColumnDimension($i)->setAutoSize(true);
                }
                $sheet->calculateColumnWidths();
                // increase columns by little padding
                $toCol = $sheet->getColumnDimension($sheet->getHighestColumn())->getColumnIndex();
                $toCol++;
                for ($i = 'A'; $i !== $toCol; $i++) {
                    $calculatedWidth = $sheet->getColumnDimension($i)->getWidth();
                    $sheet
                        ->getColumnDimension($i)
                        ->setWidth((int) $calculatedWidth * 1.1)
                        ->setAutoSize(false);
                }
                $sheet->calculateColumnWidths();
            }
            if ($args['autosize_columns'] === false) {
                $toCol = $sheet->getColumnDimension($sheet->getHighestColumn())->getColumnIndex();
                $toCol++;
                for ($i = 'A'; $i !== $toCol; $i++) {
                    $calculatedWidth = $sheet->getColumnDimension($i)->getWidth();
                    $sheet
                        ->getColumnDimension($i)
                        ->setWidth(30)
                        ->setAutoSize(false);
                }
                $sheet->calculateColumnWidths();
            }

            if ($args['auto_borders'] === true) {
                $sheet->getStyle('A1:' . ($sheet->getHighestColumn() . $sheet->getHighestRow()))->applyFromArray([
                    'borders' => [
                        'allBorders' => [
                            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                            'color' => ['rgb' => '000000']
                        ]
                    ]
                ]);
            }

            // format some data as text (long numbers)
            for ($row = 1; $row <= $sheet->getHighestRow(); $row++) {
                $toCol = $sheet->getHighestColumn();
                $toCol++;
                for ($col = 'A'; $col != $toCol; $col++) {
                    $val = $sheet->getCell($col . $row)->getValue();
                    if (strpos($val, '\'') === 0 && is_numeric(str_replace('\'', '', $val)) && strlen($val) > 10) {
                        $sheet
                            ->getCell($col . $row)
                            ->setValueExplicit(
                                str_replace('\'', '', $val),
                                \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING
                            );
                        $sheet
                            ->getStyle($col . $row)
                            ->getNumberFormat()
                            ->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_TEXT);
                    }
                }
            }

            // format some data as currency
            for ($row = 1; $row <= $sheet->getHighestRow(); $row++) {
                $toCol = $sheet->getHighestColumn();
                $toCol++;
                for ($col = 'A'; $col != $toCol; $col++) {
                    if (preg_match('/^(\d|,|\.| )+€$/', $sheet->getCell($col . $row)->getValue())) {
                        // convert in float
                        $sheet
                            ->getCell($col . $row)
                            ->setValueExplicit(
                                floatval(
                                    str_replace(
                                        '€',
                                        '',
                                        str_replace(
                                            ',',
                                            '.',
                                            str_replace(' ', '', $sheet->getCell($col . $row)->getValue())
                                        )
                                    )
                                ),
                                \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC
                            );
                        // format as euro
                        $sheet
                            ->getStyle($col . $row)
                            ->getNumberFormat()
                            ->setFormatCode('#,##0.00€');
                        $sheet
                            ->getStyle($col . $row)
                            ->getAlignment()
                            ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT);
                    }
                }
            }

            $writer = new Xlsx($spreadsheet);

            if ($args['output'] === 'download') {
                header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
                header('Content-Disposition: attachment;filename="' . $args['file'] . '"');
                header('Cache-Control: max-age=0');
                $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
            }
            try {
                $writer->save($args['output'] === 'save' ? $args['file'] : 'php://output');
            } catch (\Exception $e) {
                throw new \Exception($e->getMessage());
            }
            if ($args['output'] === 'download') {
                die();
            }
        }
        if ($args['engine'] === 'spout') {
        }
    }

    private static function prepareWriteArgs($args)
    {
        // default values
        if (!array_key_exists('file', $args) || $args['file'] === null) {
            $args['file'] = 'output-' . date('Y-m-d-H:i:s') . '.xlsx';
        }
        if (!array_key_exists('engine', $args) || $args['engine'] === null) {
            $args['engine'] = 'phpspreadsheet';
        }
        if (!array_key_exists('output', $args) || $args['output'] === null) {
            $args['output'] = 'save';
        }
        if (!array_key_exists('style_header', $args) || $args['style_header'] === null) {
            $args['style_header'] = false;
        }
        if (!array_key_exists('autosize_columns', $args) || $args['autosize_columns'] === null) {
            $args['autosize_columns'] = true;
        }
        if (!array_key_exists('auto_borders', $args) || $args['auto_borders'] === null) {
            $args['auto_borders'] = true;
        }
        if (!array_key_exists('remove_empty_cols', $args) || $args['remove_empty_cols'] === null) {
            $args['remove_empty_cols'] = false;
        }
        if (!array_key_exists('data', $args) || $args['data'] === null) {
            $args['data'] = [];
        }
        // checks
        if (self::getType($args['file']) === null) {
            throw new \Exception('unknown file extension');
        }
        if (!in_array($args['engine'], ['phpspreadsheet', 'spout'])) {
            throw new \Exception('unknown engine');
        }
        if (!in_array($args['output'], ['save', 'download'])) {
            throw new \Exception('unknown output');
        }
        // normalize data
        $args['data'] = array_values($args['data']);
        foreach ($args['data'] as $data__key => $data__value) {
            $args['data'][$data__key] = array_values($data__value);
        }
        if ($args['remove_empty_cols'] === true) {
            $empty_rows = [];
            foreach ($args['data'] as $data__value) {
                foreach ($data__value as $data__value__key => $data__value__value) {
                    if ($data__value__value === null || $data__value__value === '') {
                        if (!array_key_exists($data__value__key, $empty_rows)) {
                            $empty_rows[$data__value__key] = 0;
                        }
                        $empty_rows[$data__value__key]++;
                    }
                }
            }
            foreach ($empty_rows as $empty_rows__key => $empty_rows__value) {
                if ($empty_rows__value >= count($args['data']) - 1) {
                    foreach ($args['data'] as $data__key => $data__value) {
                        unset($args['data'][$data__key][$empty_rows__key]);
                    }
                }
            }
            foreach ($args['data'] as $data__key => $data__value) {
                $args['data'][$data__key] = array_values($data__value);
            }
        }
        return $args;
    }

    private static function getType($file)
    {
        if (strpos($file, '.xlsx') === strlen($file) - strlen('.xlsx')) {
            return 'xlsx';
        }
        return null;
    }

    private static function char2int($letters)
    {
        $num = 0;
        $arr = array_reverse(str_split($letters));
        for ($i = 0; $i < count($arr); $i++) {
            $num += (ord(strtolower($arr[$i])) - 96) * pow(26, $i);
        }
        return $num;
    }
    private static function int2char($num)
    {
        $letters = '';
        while ($num > 0) {
            $code = $num % 26 == 0 ? 26 : $num % 26;
            $letters .= chr($code + 64);
            $num = ($num - $code) / 26;
        }
        return strtoupper(strrev($letters));
    }
}
