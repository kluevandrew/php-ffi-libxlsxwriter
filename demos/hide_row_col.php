<?php

/**
 * @see http://libxlsxwriter.github.io/hide_row_col_8c-example.html
 * @noinspection PhpUnhandledExceptionInspection
 */

use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Options\RowColOptions;
use FFILibXlsxWriter\Workbook;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

$workbook = new Workbook(__DIR__ . '/output/hide_row_col.xlsx');
$worksheet = $workbook->addWorksheet();

/* Write some data. */
$worksheet->writeString([0, 3], "Some hidden columns.");
$worksheet->writeString([7, 0], "Some hidden rows.");
/* Hide all rows without data. */
$worksheet->setDefaultRow(15, true);
/* Set the height of empty rows that we want to display even if it is */
/* the default height. */
for ($row = 1; $row <= 6; $row++) {
    $worksheet->setRow($row, 15);
}
/* Columns can be hidden explicitly. This doesn't increase the file size. */
$options = new RowColOptions(1, 0, 0);
$worksheet->setColumn("G:XFD", 8.43, null, $options);

$workbook->close();
