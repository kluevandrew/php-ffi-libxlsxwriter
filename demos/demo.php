<?php

/**
 * @see http://libxlsxwriter.github.io/demo_8c-example.html
 */

use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Workbook;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

$workbook = new Workbook(__DIR__ . '/output/demo.xlsx');
$worksheet = $workbook->addWorksheet();

/* Add a format. */
$format = $workbook->addFormat();
/* Set the bold property for the format */
$format->setBold();
/* Change the column width for clarity. */
$worksheet->setColumn(0, 0, 20);
/* Write some simple text. */
$worksheet->writeString(0, 0, "Hello");
/* Text with formatting. */
$worksheet->writeString(1, 0, "World", $format);
/* Write some numbers. */
$worksheet->writeNumber(2, 0, 123);
$worksheet->writeNumber(3, 0, 123.456);
/* Insert an image. */
$worksheet->insertImage(1, 2, __DIR__ . './logo.png'); // ATTENTION: Absolute path only!

$workbook->close();
