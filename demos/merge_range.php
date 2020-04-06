<?php

/**
 * @see http://libxlsxwriter.github.io/merge_range_8c-example.html
 * @noinspection PhpUnhandledExceptionInspection
 */

use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Enums\Align;
use FFILibXlsxWriter\Enums\Border;
use FFILibXlsxWriter\Enums\Color;
use FFILibXlsxWriter\Workbook;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

$workbook = new Workbook(__DIR__ . '/output/merge_range.xlsx');
$worksheet = $workbook->addWorksheet();

$mergeFormat = $workbook->addFormat();
/* Configure a format for the merged range. */
$mergeFormat->setAlign(Align::CENTER);
$mergeFormat->setAlign(Align::VERTICAL_CENTER);
$mergeFormat->setBold();
$mergeFormat->setBgColor(Color::YELLOW);
$mergeFormat->setBorder(Border::THIN);
/* Increase the cell size of the merged cells to highlight the formatting. */
$worksheet->setColumn([1, 3], 12, null);
$worksheet->setRow(3, 30, null);
$worksheet->setRow(6, 30, null);
$worksheet->setRow(7, 30, null);
/* Merge 3 cells. */
$worksheet->mergeRange([3, 1, 3, 3], 'Merged Range', $mergeFormat);
/* Merge 3 cells over two rows. */
$worksheet->mergeRange([6, 1, 7, 3], 'Merged Range', $mergeFormat);

$workbook->close();
