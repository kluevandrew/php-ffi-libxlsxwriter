<?php
/**
 * @see http://libxlsxwriter.github.io/dates_and_times01_8c-example.html
 */
use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Workbook;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

$workbook = new Workbook(__DIR__ . '/output/dates_and_times01.xlsx');
$worksheet = $workbook->addWorksheet();

/* A number to display as a date. */
$number = 41333.5;

/* Add a format with date formatting. */
$format = $workbook->addFormat();
$format->setNumFormat("mmm d yyyy hh:mm AM/PM");
/* Widen the first column to make the text clearer. */
$worksheet->setColumn(0, 0, 20, null);
/* Write the number without formatting. */
$worksheet->writeNumber(0, 0, $number, null);  // 41333.5
/* Write the number with formatting. Note: the worksheet_write_datetime()
 * function is preferable for writing dates and times. This is for
 * demonstration purposes only.
 */
$worksheet->writeNumber(1, 0, $number, $format);   // Feb 28 2013 12:00 PM

$workbook->close();
