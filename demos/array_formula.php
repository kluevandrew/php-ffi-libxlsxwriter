<?php
/**
 * @see http://libxlsxwriter.github.io/array_formula_8c-example.html
 */
use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Workbook;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

$workbook = new Workbook(__DIR__ . '/output/array_formula.xlsx');
$worksheet = $workbook->addWorksheet();

/* Write some data for the formulas. */
$worksheet->writeNumber(0, 1, 500, null);
$worksheet->writeNumber(1, 1, 10, null);
$worksheet->writeNumber(4, 1, 1, null);
$worksheet->writeNumber(5, 1, 2, null);
$worksheet->writeNumber(6, 1, 3, null);
$worksheet->writeNumber(0, 2, 300, null);
$worksheet->writeNumber(1, 2, 15, null);
$worksheet->writeNumber(4, 2, 20234, null);
$worksheet->writeNumber(5, 2, 21003, null);
$worksheet->writeNumber(6, 2, 10000, null);
/* Write an array formula that returns a single value. */
$worksheet->writeArrayFormula(0, 0, 0, 0, "{=SUM(B1:C1*B2:C2)}", null);
/* Similar to above but using the RANGE macro. */
list($firstRow, $firstCol, $lastRow, $lastCol) = FFILibXlsxWriter::range('A2:A2');
$worksheet->writeArrayFormula($firstRow, $firstCol, $lastRow, $lastCol, "{=SUM(B1:C1*B2:C2)}", null);
/* Write an array formula that returns a range of values. */
$worksheet->writeArrayFormula(4, 0, 6, 0, "{=TREND(C5:C7,B5:B7)}", null);

$workbook->close();
