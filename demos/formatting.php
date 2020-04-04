<?php
/**
 * @see http://libxlsxwriter.github.io/tutorial02.html
 */
use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Workbook;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

$workbook = new Workbook(__DIR__ . '/output/formatting.xlsx');
$worksheet = $workbook->addWorksheet('First worksheet');

$row = 0;
$col = 0;
$i = 0;

$expenses = [
    ["item" => "Rent", "cost" => 1000],
    ["item" => "Gas", "cost" => 100],
    ["item" => "Food", "cost" => 300],
    ["item" => "Gym", "cost" => 50],
];

$bold = $workbook->addFormat();
$bold->setBold();

/* Add a number format for cells with money. */
$money = $workbook->addFormat();
$money->setNumFormat('$#,##0');

$worksheet->writeString( $row, $col,     "Item", $bold);
$worksheet->writeString( $row, $col + 1, "Cost", $bold);

/* Iterate over the data and write it out element by element. */
for ($i = 0; $i < 4; $i++) {
    $row = $i + 1;
    $worksheet->writeString($row, $col, $expenses[$i]['item'], null);
    $worksheet->writeNumber($row, $col + 1, $expenses[$i]['cost'], $money);
}
/* Write a total using a formula. */
$worksheet->writeString($row, $col, "Total", $bold);
$worksheet->writeFormula($row, $col + 1, "=SUM(B1:B4)", $money);

$workbook->close();
