<?php

/**
 * @see http://libxlsxwriter.github.io/format_font_8c-example.html
 */

use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Workbook;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

/* Create a new workbook. */
$workbook = new Workbook(__DIR__ . '/output/format_font.xlsx');
/* Add a worksheet. */
$worksheet = $workbook->addWorksheet();
/* Widen the first column to make the text clearer. */
$worksheet->setColumn(0, 0, 20);
/* Add some formats. */
$format1 = $workbook->addFormat();
$format2 = $workbook->addFormat();
$format3 = $workbook->addFormat();
/* Set the bold property for format 1. */
$format1->setBold();
/* Set the italic property for format 2. */
$format2->setItalic();
/* Set the bold and italic properties for format 3. */
$format3->setBold();
$format3->setItalic();
/* Write some formatted strings. */
$worksheet->writeString(0, 0, "This is bold", $format1);
$worksheet->writeString(1, 0, "This is italic", $format2);
$worksheet->writeString(2, 0, "Bold and italic", $format3);
/* Close the workbook, save the file and free any memory. */
$workbook->close();
