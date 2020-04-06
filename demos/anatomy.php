<?php

/**
 * @see http://libxlsxwriter.github.io/anatomy_8c-example.html
 * @noinspection PhpUnhandledExceptionInspection
 */

use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Workbook;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

/* Create a new workbook. */
$workbook = new Workbook(__DIR__ . '/output/anatomy.xlsx');
/* Add a worksheet with a user defined sheet name. */
$worksheet1 = $workbook->addWorksheet("Demo");
/* Add a worksheet with Excel's default sheet name: Sheet2. */
$worksheet2 = $workbook->addWorksheet();
/* Add some cell formats. */
$myformat1 = $workbook->addFormat();
$myformat2 = $workbook->addFormat();
/* Set the bold property for the first format. */
$myformat1->setBold();
/* Set a number format for the second format. */
$myformat2->setNumFormat("$#,##0.00");
/* Widen the first column to make the text clearer. */
$worksheet1->setColumn([0, 0], 20);
/* Write some unformatted data. */
$worksheet1->writeString([0, 0], "Peach");
$worksheet1->writeString([1, 0], "Plum");
/* Write formatted data. */
$worksheet1->writeString([2, 0], "Pear", $myformat1);
/* Formats can be reused. */
$worksheet1->writeString([3, 0], "Persimmon", $myformat1);
/* Write some numbers. */
$worksheet1->writeNumber([5, 0], 123);
$worksheet1->writeNumber([6, 0], 4567.555, $myformat2);
/* Write to the second worksheet. */
$worksheet2->writeString([0, 0], "Some text", $myformat1);

$workbook->close();
