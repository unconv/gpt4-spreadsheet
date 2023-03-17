<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Color;

function parseSpreadsheetContent($content)
{
    $spreadsheet = new Spreadsheet();
    $worksheet = $spreadsheet->getActiveSheet();

    $rows = explode("\n", $content);

    foreach ($rows as $rowIndex => $rowContent) {
        $columns = explode("|", $rowContent);

        foreach ($columns as $colIndex => $cellContent) {
            $cell = $worksheet->getCellByColumnAndRow($colIndex + 1, $rowIndex + 1);

            preg_match_all('/\[(.*?)\]/', $cellContent, $matches);

            foreach ($matches[1] as $match) {
                $parts = explode(': ', $match);

                switch ($parts[0]) {
                    case 'style':
                        if ($parts[1] == 'bold') {
                            $worksheet->getStyle($cell->getCoordinate())->getFont()->setBold(true);
                        }
                        break;

                    case 'formula':
                        $cell->setValue($parts[1]);
                        break;

                    case 'bg':
                        $color = ($parts[1] == 'yellow') ? Color::COLOR_YELLOW : Color::COLOR_WHITE;
                        $worksheet->getStyle($cell->getCoordinate())->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB($color);
                        break;
                }
            }

            $cleanContent = preg_replace('/\[(.*?)\]/', '', $cellContent);
            if (trim($cleanContent) !== '') {
                $cell->setValue($cleanContent);
            }
        }
    }

    return $spreadsheet;
}

$content = file_get_contents('csv-spreadsheet.csv');
$spreadsheet = parseSpreadsheetContent($content);

$writer = new Xlsx($spreadsheet);
$writer->save('MathTest3.xlsx');
