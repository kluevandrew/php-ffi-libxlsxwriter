<?php

/**
 * @see http://libxlsxwriter.github.io/images_8c-example.html
 * @noinspection PhpUnhandledExceptionInspection
 */

use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Options\ImageOptions;
use FFILibXlsxWriter\Workbook;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

/* Create a new workbook and add a worksheet. */
$workbook = new Workbook(__DIR__ . '/output/images.xlsx');
$worksheet = $workbook->addWorksheet();

/* Change some of the column widths for clarity. */
$worksheet->setColumn('A:A', 30);

/* Insert an image. */
$worksheet->writeString('A2', 'Insert an image in a cell:');
$worksheet->insertImage('B2', __DIR__ . '/logo.png');

/* Insert an image offset in the cell. */
$worksheet->writeString('A12', 'Insert an offset image:');
$options1 = new ImageOptions();
$options1->x_offset = 15;
$options1->y_offset = 10;
$worksheet->insertImage('B12', __DIR__ . '/logo.png', $options1);

/* Insert an image with scaling. */
$worksheet->writeString('A22', 'Insert a scaled image:');
$options2 = new ImageOptions();
$options2->x_scale = 0.5;
$options2->y_scale = 0.5;
$worksheet->insertImage('B22', __DIR__ . '/logo.png', $options2);

/* Insert an image with a hyperlink. */
$worksheet->writeString('A32', 'Insert an image with a hyperlink:');
$options3 = new ImageOptions();
$options3->url = 'https://github.com/kluevandrew/php-ffi-libxlsxwriter';
$worksheet->insertImage('B32', __DIR__ . '/logo.png', $options3);
$workbook->close();
