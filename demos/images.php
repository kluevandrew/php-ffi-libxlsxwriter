<?php

/**
 * @see http://libxlsxwriter.github.io/images_8c-example.html
 */

use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Structs\ImageOptions;
use FFILibXlsxWriter\Workbook;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

/* Create a new workbook and add a worksheet. */
$workbook = new Workbook(__DIR__ . '/output/images.xlsx');
$worksheet = $workbook->addWorksheet();

/* Change some of the column widths for clarity. */
list($firstCol, $lastCol) = FFILibXlsxWriter::cols('A:A');
$worksheet->setColumn($firstCol, $lastCol, 30);
/* Insert an image. */
list($row, $col) = FFILibXlsxWriter::cell('A2');
$worksheet->writeString($row, $col, 'Insert an image in a cell:');
list($row, $col) = FFILibXlsxWriter::cell('B2');
$worksheet->insertImage($row, $col, __DIR__ . '/logo.png');
/* Insert an image offset in the cell. */
list($row, $col) = FFILibXlsxWriter::cell("A12");
$worksheet->writeString($row, $col, "Insert an offset image:");
$options1 = new ImageOptions();
$options1->x_offset = 15;
$options1->y_offset = 10;
list($row, $col) = FFILibXlsxWriter::cell("B12");
$worksheet->insertImageOpt($row, $col, __DIR__ . '/logo.png', $options1);
/* Insert an image with scaling. */
list($row, $col) = FFILibXlsxWriter::cell("A22");
$worksheet->writeString($row, $col, "Insert a scaled image:");
$options2 = new ImageOptions();
$options2->x_scale = 0.5;
$options2->y_scale = 0.5;
list($row, $col) = FFILibXlsxWriter::cell("B22");
$worksheet->insertImageOpt($row, $col, __DIR__ . '/logo.png', $options2);
/* Insert an image with a hyperlink. */
list($row, $col) = FFILibXlsxWriter::cell("A32");
$worksheet->writeString($row, $col, "Insert an image with a hyperlink:");
$options3 = new ImageOptions();
$options3->url = "https://github.com/kluevandrew/php-ffi-libxlsxwriter";
list($row, $col) = FFILibXlsxWriter::cell("B32");
$worksheet->insertImageOpt($row, $col, __DIR__ . '/logo.png', $options3);
$workbook->close();
