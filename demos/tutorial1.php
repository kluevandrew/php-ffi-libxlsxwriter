<?php

/**
 * @see http://libxlsxwriter.github.io/tutorial1_8c-example.html
 * @noinspection PhpUnhandledExceptionInspection
 */

use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Workbook;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

$workbook = new Workbook(__DIR__ . '/output/tutorial1.xlsx');
$worksheet = $workbook->addWorksheet('First worksheet');

$col = 0;
$expenses = [
    ["item" => "Rent", "cost" => 1000],
    ["item" => "Gas", "cost" => 100],
    ["item" => "Food", "cost" => 300],
    ["item" => "Gym", "cost" => 50],
];

/* Iterate over the data and write it out element by element. */
for ($row = 0; $row < 4; $row++) {
    $worksheet->writeString([$row, $col], $expenses[$row]['item'], null);
    $worksheet->writeNumber([$row, $col + 1], $expenses[$row]['cost'], null);
}
/* Write a total using a formula. */
$worksheet->writeString([$row, $col], "Total", null);
$worksheet->writeFormula([$row, $col + 1], "=SUM(B1:B4)", null);

$workbook->close();
