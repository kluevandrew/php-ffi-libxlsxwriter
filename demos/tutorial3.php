<?php

/**
 * @see http://libxlsxwriter.github.io/tutorial3_8c-example.html
 */

use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Structs\DateTime;
use FFILibXlsxWriter\Workbook;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

$workbook = new Workbook(__DIR__ . '/output/tutorial3.xlsx');
$worksheet = $workbook->addWorksheet('First worksheet');

$row = 0;
$col = 0;
$i = 0;

$expenses = [
    ["item" => "Rent", "datetime" => new DateTime(2020, 1, 13), "cost" => 1000],
    ["item" => "Gas", "datetime" => new DateTime(2020, 1, 14), "cost" => 100],
    ["item" => "Food", "datetime" => new DateTime(2020, 1, 16), "cost" => 300],
    ["item" => "Gym", "datetime" => new DateTime(2020, 1, 20), "cost" => 50],
];

$bold = $workbook->addFormat();
$bold->setBold();

/* Add a number format for cells with money. */
$money = $workbook->addFormat();
$money->setNumFormat('$#,##0');

/* Add a number format for cells with money. */
$dateFormat = $workbook->addFormat();
$dateFormat->setNumFormat('mmmm d yyyy');

$worksheet->writeString($row, $col, "Item", $bold);
$worksheet->writeString($row, $col + 1, "Cost", $bold);

/* Iterate over the data and write it out element by element. */
for ($i = 0; $i < 4; $i++) {
    $row = $i + 1;
    $worksheet->writeString($row, $col, $expenses[$i]['item'], null);
    $worksheet->writeDatetime($row, $col + 1, $expenses[$i]['datetime'], $dateFormat);
    $worksheet->writeNumber($row, $col + 2, $expenses[$i]['cost'], $money);
}
/* Write a total using a formula. */
$worksheet->writeString($row + 1, $col, "Total", $bold);
$worksheet->writeFormula($row + 1, $col + 2, "=SUM(C2:C5)", $money);

$workbook->close();
