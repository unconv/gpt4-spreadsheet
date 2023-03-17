<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;

// Create a new spreadsheet
$spreadsheet = new Spreadsheet();
$worksheet = $spreadsheet->getActiveSheet();

// Set the student name label
$worksheet->setCellValue('A1', 'Student Name:');
$worksheet->setCellValue('A2', '');

// Add math questions
$questions = [
    '1. What is 36 divided by 4?',
    '2. Calculate the area of a rectangle with a length of 7 units and a width of 5 units.',
    '3. What is the least common multiple of 4 and 6?',
    '4. Simplify the fraction 20/28.',
    '5. Solve the equation: 2x + 6 = 16',
    '6. Find the perimeter of a square with a side length of 8 units.',
    '7. What is the value of the expression 3^3?',
    '8. Convert 0.25 to a fraction.',
    '9. How many sides does a hexagon have?',
    '10. Write the number 256 in expanded form.'
];

$currentRow = 4;

for ($i = 0; $i < count($questions); $i++) {
    $worksheet->setCellValue("A{$currentRow}", $questions[$i]);
    $worksheet->getStyle("A{$currentRow}")->getFont()->setBold(true);

    // Add 4 empty rows for the answers
    for ($j = 1; $j <= 4; $j++) {
        $currentRow++;
        $worksheet->setCellValue("A{$currentRow}", '');
        $worksheet->setCellValue("B{$currentRow}", '');
    }

    // Add points cell on the last answer row
    $worksheet->getStyle("B{$currentRow}")->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB(Color::COLOR_YELLOW);
    $currentRow++;
}

// Calculate the total points earned
$worksheet->setCellValue("B{$currentRow}", '=SUM(B4:B13)');

// Calculate the maximum possible points
$currentRow++;
$worksheet->setCellValue("B{$currentRow}", '=5*10');

// Calculate the grade
$currentRow++;
$worksheet->setCellValue("B{$currentRow}", '=IF(B14/B15>=0.9,"A",IF(B14/B15>=0.8,"B",IF(B14/B15>=0.7,"C",IF(B14/B15>=0.6,"D","F"))))');

// Save the spreadsheet as an XLSX file
$writer = new Xlsx($spreadsheet);
$writer->save('MathTest2.xlsx');

?>
