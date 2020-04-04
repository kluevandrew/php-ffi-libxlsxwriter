<?php
/**
 * @see http://libxlsxwriter.github.io/dates_and_times02_8c-example.html
 */

use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Workbook;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

$workbook = new Workbook(__DIR__ . '/output/dates_and_times02.xlsx');
$worksheet = $workbook->addWorksheet();

/* A datetime to display. */
$datetime = new \FFILibXlsxWriter\Structs\DateTime(2013, 2, 28, 12, 0, 0.0);
/* Add a format with date formatting. */
$format = $workbook->addFormat();
$format->setNumFormat("mmm d yyyy hh:mm AM/PM");
/* Widen the first column to make the text clearer. */
$worksheet->setColumn(0, 0, 20, null);
/* Write the datetime without formatting. */
$worksheet->writeDatetime(0, 0, $datetime, null);  // 41333.5
/* Write the datetime with formatting. */
$worksheet->writeDatetime(1, 0, $datetime, $format);  // Feb 28 2013 12:00 PM

$workbook->close();
