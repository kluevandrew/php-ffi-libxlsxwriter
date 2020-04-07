<?php

/**
 * @see http://libxlsxwriter.github.io/doc_custom_properties_8c-example.html
 * @noinspection PhpUnhandledExceptionInspection
 */

use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Structs\DateTime as XlsxDateTime;
use FFILibXlsxWriter\Workbook;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

$workbook = new Workbook(__DIR__ . '/output/doc_custom_properties.xlsx');
$worksheet = $workbook->addWorksheet();
$datetime = new XlsxDateTime(2016, 12, 12, 0, 0, 0.0);
/* Set some custom document properties in the workbook. */
$workbook->setCustomPropertyString("Checked by", "Eve");
$workbook->setCustomPropertyDateTime("Date completed", $datetime);
$workbook->setCustomPropertyNumber("Document number", 12345);
$workbook->setCustomPropertyNumber("Reference number", 1.2345);
$workbook->setCustomPropertyBoolean("Has Review", true);
$workbook->setCustomPropertyBoolean("Signed off", false);
/* Add some text to the file. */
$worksheet->setColumn('A:A', 50);
$worksheet->writeString('A1', 'Select \'Workbook Properties\' to see properties.');

$workbook->close();
