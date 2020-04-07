<?php

/**
 * @see http://libxlsxwriter.github.io/hide_sheet_8c-example.html
 * @noinspection PhpUnhandledExceptionInspection
 */

use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Workbook;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

$workbook = new Workbook(__DIR__ . '/output/hide_sheet.xlsx');

$worksheet1 = $workbook->addWorksheet();
$worksheet2 = $workbook->addWorksheet();
$worksheet3 = $workbook->addWorksheet();

/* Hide Sheet2. It won't be visible until it is unhidden in Excel. */
$worksheet2->hide();
$worksheet1->writeString('A1', 'Sheet2 is hidden');
$worksheet2->writeString('A1', 'Now it\'s my turn to find you!');
$worksheet3->writeString('A1', "Sheet2 is hidden");

/* Make the first column wider to make the text clearer. */
$worksheet1->setColumn('A:A', 30);
$worksheet2->setColumn('A:A', 30);
$worksheet3->setColumn('A:A', 30);

$workbook->close();
