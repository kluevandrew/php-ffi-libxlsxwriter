<?php

/**
 * @see http://libxlsxwriter.github.io/format_num_format_8c-example.html
 */

use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Workbook;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

/* Create a new workbook and add a worksheet. */
$workbook = new Workbook(__DIR__ . '/output/format_num_format.xlsx');
$worksheet = $workbook->addWorksheet();
/* Widen the first column to make the text clearer. */
$worksheet->setColumn(0, 0, 30);
/* Add some formats. */
$format01 = $workbook->addFormat();
$format02 = $workbook->addFormat();
$format03 = $workbook->addFormat();
$format04 = $workbook->addFormat();
$format05 = $workbook->addFormat();
$format06 = $workbook->addFormat();
$format07 = $workbook->addFormat();
$format08 = $workbook->addFormat();
$format09 = $workbook->addFormat();
$format10 = $workbook->addFormat();
$format11 = $workbook->addFormat();
/* Set some example number formats. */
$format01->setNumFormat('0.000');
$format02->setNumFormat('#,##0');
$format03->setNumFormat('#,##0.00');
$format04->setNumFormat('0.00');
$format05->setNumFormat('mm/dd/yy');
$format06->setNumFormat('mmm d yyyy');
$format07->setNumFormat('d mmmm yyyy');
$format08->setNumFormat('dd/mm/yyyy hh:mm AM/PM');
$format09->setNumFormat('0 "dollar and" .00 "cents"');
/* Write data using the formats. */
$worksheet->writeNumber(0, 0, 3.1415926);      // 3.1415926
$worksheet->writeNumber(1, 0, 3.1415926, $format01);  // 3.142
$worksheet->writeNumber(2, 0, 1234.56, $format02);  // 1,235
$worksheet->writeNumber(3, 0, 1234.56, $format03);  // 1,234.56
$worksheet->writeNumber(4, 0, 49.99, $format04);  // 49.99
$worksheet->writeNumber(5, 0, 36892.521, $format05);  // 01/01/01
$worksheet->writeNumber(6, 0, 36892.521, $format06);  // Jan 1 2001
$worksheet->writeNumber(7, 0, 36892.521, $format07);  // 1 January 2001
$worksheet->writeNumber(8, 0, 36892.521, $format08);  // 01/01/2001 12:30 AM
$worksheet->writeNumber(9, 0, 1.87, $format09);  // 1 dollar and .87 cents
/* Show limited conditional number formats. */
$format10->setNumFormat("[Green]General;[Red]-General;General");
$worksheet->writeNumber(10, 0, 123, $format10);  // > 0 Green
$worksheet->writeNumber(11, 0, -45, $format10);  // < 0 Red
$worksheet->writeNumber(12, 0, 0, $format10);  // = 0 Default color
/* Format a Zip code. */
$format11->setNumFormat("00000");
$worksheet->writeNumber(13, 0, 1209, $format11);

/* Close the workbook, save the file and free any memory. */
$workbook->close();
