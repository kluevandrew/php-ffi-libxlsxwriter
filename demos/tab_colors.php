<?php

/**
 * @see http://libxlsxwriter.github.io/tab_colors_8c-example.html
 */

use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Structs\Color;
use FFILibXlsxWriter\Workbook;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

$workbook = new Workbook(__DIR__ . '/output/tab_colors.xlsx');
/* Set up some worksheets. */
$worksheet1 = $workbook->addWorksheet();
$worksheet2 = $workbook->addWorksheet();
$worksheet3 = $workbook->addWorksheet();
$worksheet4 = $workbook->addWorksheet();
/* Set the tab colors. */
$worksheet1->setTabColor(Color::RED);
$worksheet2->setTabColor(Color::GREEN);
$worksheet3->setTabColor(0xFF9900); /* Orange. */
/* worksheet4 will have the default color. */
$worksheet4->writeString(0, 0, "Hello");

$workbook->close();
