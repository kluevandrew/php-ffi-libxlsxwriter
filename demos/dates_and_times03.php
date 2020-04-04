<?php
/**
 * @see http://libxlsxwriter.github.io/dates_and_times03_8c-example.html
 */

use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Structs\Align;
use FFILibXlsxWriter\Workbook;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

/* A datetime to display. */
$datetime = new \FFILibXlsxWriter\Structs\DateTime(2013, 1, 23, 12, 30, 5.123);
$row = 0;
$col = 0;
$i = 0;
/* Examples date and time formats. In the output file compare how changing
 * the format strings changes the appearance of the date.
 */
$date_formats = [
    "dd/mm/yy",
    "mm/dd/yy",
    "dd m yy",
    "d mm yy",
    "d mmm yy",
    "d mmmm yy",
    "d mmmm yyy",
    "d mmmm yyyy",
    "dd/mm/yy hh:mm",
    "dd/mm/yy hh:mm:ss",
    "dd/mm/yy hh:mm:ss.000",
    "hh:mm",
    "hh:mm:ss",
    "hh:mm:ss.000",
];
/* Create a new workbook and add a worksheet. */
$workbook = new Workbook(__DIR__ . '/output/dates_and_times03.xlsx');
$worksheet = $workbook->addWorksheet();
/* Add a bold format. */
$bold = $workbook->addFormat();
$bold->setBold();
/* Write the column headers. */
$worksheet->writeString($row, $col, "Formatted date", $bold);
$worksheet->writeString($row, $col + 1, "Format", $bold);
/* Widen the first column to make the text clearer. */
$worksheet->setColumn( 0, 1, 20, null);
/* Write the same date and time using each of the above formats. */
for ($i = 0; $i < 14; $i++) {
    $row++;
    /* Create a format for the date or time.*/
    $format = $workbook->addFormat();
    $format->setNumFormat($date_formats[$i]);
    $format->setAlign(Align::LEFT);
    /* Write the datetime with each format. */
    $worksheet->writeDatetime( $row, $col, $datetime, $format);
    /* Also write the format string for comparison. */
    $worksheet->writeString($row, $col + 1, $date_formats[$i], null);
}
$workbook->close();
