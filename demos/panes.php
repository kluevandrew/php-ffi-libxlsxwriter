<?php
/**
 * @see http://libxlsxwriter.github.io/panes_8c-example.html
 */
use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Structs\Align;
use FFILibXlsxWriter\Structs\Border;
use FFILibXlsxWriter\Workbook;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

$workbook = new Workbook(__DIR__ . '/output/panes.xlsx');

$worksheet1 = $workbook->addWorksheet("Panes 1");
$worksheet2 = $workbook->addWorksheet("Panes 2");
$worksheet3 = $workbook->addWorksheet("Panes 3");
$worksheet4 = $workbook->addWorksheet("Panes 4");

/* Set up some formatting and text to highlight the panes. */
$header = $workbook->addFormat();;
$header->setAlign(Align::CENTER);
$header->setAlign(Align::VERTICAL_CENTER);
$header->setFgColor(0xD7E4BC);
$header->setBold();
$header->setBorder(Border::THIN);
$center = $workbook->addFormat();;
$center->setAlign(Align::CENTER);
/*
 * Example 1. Freeze pane on the top row.
 */
$worksheet1->freezePanes(1, 0);
/* Some sheet formatting. */
$worksheet1->setColumn(0, 8, 16, null);
$worksheet1->setRow(0, 20, null);
$worksheet1->setSelection(4, 3, 4, 3);
/* Some worksheet text to demonstrate scrolling. */
for ($col = 0; $col < 9; $col++) {
    $worksheet1->writeString(0, $col, "Scroll down", $header);
}
for ($row = 1; $row < 100; $row++) {
    for ($col = 0; $col < 9; $col++) {
        $worksheet1->writeNumber($row, $col, $row + 1, $center);
    }
}
/*
 * Example 2. Freeze pane on the left column.
 */
$worksheet2->freezePanes( 0, 1);
/* Some sheet formatting. */
$worksheet2->setColumn(0, 0, 16, NULL);
$worksheet2->setSelection(4, 3, 4, 3);
/* Some worksheet text to demonstrate scrolling. */
for ($row = 0; $row < 50; $row++) {
    $worksheet2->writeString($row, 0, "Scroll right", $header);
    for ($col = 1; $col < 26; $col++) {
        $worksheet2->writeNumber($row, $col, $col, $center);
    }
}
/*
 * Example 3. Freeze pane on the top row and left column.
 */
$worksheet3->freezePanes( 1, 1);
/* Some sheet formatting. */
$worksheet3->setColumn(0, 25, 16, NULL);
$worksheet3->setRow(0, 20, NULL);
$worksheet3->writeString(0, 0, "", $header);
$worksheet3->setSelection(4, 3, 4, 3);
/* Some worksheet text to demonstrate scrolling. */
for ($col = 1; $col < 26; $col++) {
    $worksheet3->writeString(0, $col, "Scroll down", $header);
}
for ($row = 1; $row < 50; $row++) {
    $worksheet3->writeString($row, 0, "Scroll right", $header);
    for ($col = 1; $col < 26; $col++) {
        $worksheet3->writeNumber($row, $col, $col, $center);
    }
}
/*
 * Example 4. Split pane on the top row and left column.
 *
 * The divisions must be specified in terms of row and column dimensions.
 * The default row height is 15 and the default column width is 8.43
 */
$worksheet4->splitPanes(15, 8.43);
/* Some sheet formatting. */
/* Some worksheet text to demonstrate scrolling. */
for ($col = 1; $col < 26; $col++) {
    $worksheet4->writeString(0, $col, "Scroll", $center);
}
for ($row = 1; $row < 50; $row++) {
    $worksheet4->writeString($row, 0, "Scroll", $center);
    for ($col = 1; $col < 26; $col++) {
        $worksheet4->writeNumber($row, $col, $col, $center);
    }
}

$workbook->close();
