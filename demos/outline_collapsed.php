<?php

/**
 * @see http://libxlsxwriter.github.io/outline_collapsed_8c-example.html
 * @noinspection PhpUnhandledExceptionInspection
 * @noinspection DuplicatedCode
 */

use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Options\RowColOptions;
use FFILibXlsxWriter\Style\Format;
use FFILibXlsxWriter\Workbook;
use FFILibXlsxWriter\Worksheet;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

/**
 * This function will generate the same data and sub-totals on each worksheet.
 * Used in the examples 1-4.
 * @param Worksheet $worksheet
 * @param Format $bold
 */
$create_row_example_data = function (Worksheet $worksheet, Format $bold) {
    /* Set the column width for clarity. */
    $worksheet->setColumn('A:A', 20);
    /* Add data and formulas to the worksheet. */
    $worksheet->writeString('A1', "Region", $bold);
    $worksheet->writeString('A2', "North");
    $worksheet->writeString('A3', "North");
    $worksheet->writeString('A4', "North");
    $worksheet->writeString('A5', "North");
    $worksheet->writeString('A6', "North Total", $bold);
    $worksheet->writeString('B1', "Sales", $bold);
    $worksheet->writeNumber('B2', 1000);
    $worksheet->writeNumber('B3', 1200);
    $worksheet->writeNumber('B4', 900);
    $worksheet->writeNumber('B5', 1200);
    $worksheet->writeFormula('B6', "=SUBTOTAL(9,B2:B5)", $bold);
    $worksheet->writeString('A7', "South");
    $worksheet->writeString('A8', "South");
    $worksheet->writeString('A9', "South");
    $worksheet->writeString('A10', "South");
    $worksheet->writeString('A11', "South Total", $bold);
    $worksheet->writeNumber('B7', 400);
    $worksheet->writeNumber('B8', 600);
    $worksheet->writeNumber('B9', 500);
    $worksheet->writeNumber('B10', 600);
    $worksheet->writeFormula('B11', "=SUBTOTAL(9,B7:B10)", $bold);
    $worksheet->writeString('A12', "Grand Total", $bold);
    $worksheet->writeFormula('B12', "=SUBTOTAL(9,B2:B10)", $bold);
};

/**
 * This function will generate the same data and sub-totals on each worksheet.
 * Used in the examples 5-6.
 * @param Worksheet $worksheet
 * @param Format $bold
 */
$create_col_example_data = function (Worksheet $worksheet, Format $bold) {
    /* Add data and formulas to the worksheet. */
    $worksheet->writeString('A1', "Month");
    $worksheet->writeString('B1', "Jan");
    $worksheet->writeString('C1', "Feb");
    $worksheet->writeString('D1', "Mar");
    $worksheet->writeString('E1', "Apr");
    $worksheet->writeString('F1', "May");
    $worksheet->writeString('G1', "Jun");
    $worksheet->writeString('H1', "Total");
    $worksheet->writeString('A2', "North");
    $worksheet->writeNumber('B2', 50);
    $worksheet->writeNumber('C2', 20);
    $worksheet->writeNumber('D2', 15);
    $worksheet->writeNumber('E2', 25);
    $worksheet->writeNumber('F2', 65);
    $worksheet->writeNumber('G2', 80);
    $worksheet->writeFormula('H2', "=SUM(B2:G2)");
    $worksheet->writeString('A3', "South");
    $worksheet->writeNumber('B3', 10);
    $worksheet->writeNumber('C3', 20);
    $worksheet->writeNumber('D3', 30);
    $worksheet->writeNumber('E3', 50);
    $worksheet->writeNumber('F3', 50);
    $worksheet->writeNumber('G3', 50);
    $worksheet->writeFormula('H3', "=SUM(B3:G3)");
    $worksheet->writeString('A4', "East");
    $worksheet->writeNumber('B4', 45);
    $worksheet->writeNumber('C4', 75);
    $worksheet->writeNumber('D4', 50);
    $worksheet->writeNumber('E4', 15);
    $worksheet->writeNumber('F4', 75);
    $worksheet->writeNumber('G4', 100);
    $worksheet->writeFormula('H4', "=SUM(B4:G4)");
    $worksheet->writeString('A5', "West");
    $worksheet->writeNumber('B5', 15);
    $worksheet->writeNumber('C5', 15);
    $worksheet->writeNumber('D5', 55);
    $worksheet->writeNumber('E5', 35);
    $worksheet->writeNumber('F5', 20);
    $worksheet->writeNumber('G5', 50);
    $worksheet->writeFormula('H5', "=SUM(B5:G5)");
    $worksheet->writeFormula('H6', "=SUM(H2:H5)", $bold);
};

$workbook = new Workbook(__DIR__ . '/output/outline_collapsed.xlsx');
$worksheet1 = $workbook->addWorksheet("Outlined Rows");
$worksheet2 = $workbook->addWorksheet("Collapsed Rows 1");
$worksheet3 = $workbook->addWorksheet("Collapsed Rows 2");
$worksheet4 = $workbook->addWorksheet("Collapsed Rows 3");
$worksheet5 = $workbook->addWorksheet("Outline Columns");
$worksheet6 = $workbook->addWorksheet("Collapsed Columns");
$bold = $workbook->addFormat();
$bold->setBold();

/*
* Example 1: Create a worksheet with outlined rows. It also includes
* SUBTOTAL() functions so that it looks like the type of automatic
* outlines that are generated when you use the 'Sub Totals' option.
*
* For outlines the important parameters are 'hidden' and 'level'. Rows
* with the same 'level' are grouped together. The group will be collapsed
* if 'hidden' is non-zero.
*/
/* The option structs with the outline level set. */
$options1 = new RowColOptions(0, 2, 0);
$options2 = new RowColOptions(0, 1, 0);
/* Set the row outline properties set. */
$worksheet1->setRow(1, RowColOptions::DEF_ROW_HEIGHT, null, $options1);
$worksheet1->setRow(2, RowColOptions::DEF_ROW_HEIGHT, null, $options1);
$worksheet1->setRow(3, RowColOptions::DEF_ROW_HEIGHT, null, $options1);
$worksheet1->setRow(4, RowColOptions::DEF_ROW_HEIGHT, null, $options1);
$worksheet1->setRow(5, RowColOptions::DEF_ROW_HEIGHT, null, $options2);
$worksheet1->setRow(6, RowColOptions::DEF_ROW_HEIGHT, null, $options1);
$worksheet1->setRow(7, RowColOptions::DEF_ROW_HEIGHT, null, $options1);
$worksheet1->setRow(8, RowColOptions::DEF_ROW_HEIGHT, null, $options1);
$worksheet1->setRow(9, RowColOptions::DEF_ROW_HEIGHT, null, $options1);
$worksheet1->setRow(10, RowColOptions::DEF_ROW_HEIGHT, null, $options2);
/* Write the sub-total data that is common to the row examples. */
$create_row_example_data($worksheet1, $bold);

/*
* Example 2: Create a worksheet with collapsed outlined rows.
* This is the same as the example 1  except that the all rows are collapsed.
*/
/* The option structs with the outline properties set. */
$options3 = new RowColOptions(1, 2, 0);
$options4 = new RowColOptions(1, 1, 0);
$options5 = new RowColOptions(0, 0, 1);
/* Set the row options with the outline level. */
$worksheet2->setRow(1, RowColOptions::DEF_ROW_HEIGHT, null, $options3);
$worksheet2->setRow(2, RowColOptions::DEF_ROW_HEIGHT, null, $options3);
$worksheet2->setRow(3, RowColOptions::DEF_ROW_HEIGHT, null, $options3);
$worksheet2->setRow(4, RowColOptions::DEF_ROW_HEIGHT, null, $options3);
$worksheet2->setRow(5, RowColOptions::DEF_ROW_HEIGHT, null, $options4);
$worksheet2->setRow(6, RowColOptions::DEF_ROW_HEIGHT, null, $options3);
$worksheet2->setRow(7, RowColOptions::DEF_ROW_HEIGHT, null, $options3);
$worksheet2->setRow(8, RowColOptions::DEF_ROW_HEIGHT, null, $options3);
$worksheet2->setRow(9, RowColOptions::DEF_ROW_HEIGHT, null, $options3);
$worksheet2->setRow(10, RowColOptions::DEF_ROW_HEIGHT, null, $options4);
$worksheet2->setRow(11, RowColOptions::DEF_ROW_HEIGHT, null, $options5);
/* Write the sub-total data that is common to the row examples. */
$create_row_example_data($worksheet2, $bold);

/*
 * Example 3: Create a worksheet with collapsed outlined rows. Same as the
 * example 1 except that the two sub-totals are collapsed.
 */
$options6 = new RowColOptions(1, 2, 0);
$options7 = new RowColOptions(0, 1, 1);
/* Set the row options with the outline level. */
$worksheet3->setRow(1, RowColOptions::DEF_ROW_HEIGHT, null, $options6);
$worksheet3->setRow(2, RowColOptions::DEF_ROW_HEIGHT, null, $options6);
$worksheet3->setRow(3, RowColOptions::DEF_ROW_HEIGHT, null, $options6);
$worksheet3->setRow(4, RowColOptions::DEF_ROW_HEIGHT, null, $options6);
$worksheet3->setRow(5, RowColOptions::DEF_ROW_HEIGHT, null, $options7);
$worksheet3->setRow(6, RowColOptions::DEF_ROW_HEIGHT, null, $options6);
$worksheet3->setRow(7, RowColOptions::DEF_ROW_HEIGHT, null, $options6);
$worksheet3->setRow(8, RowColOptions::DEF_ROW_HEIGHT, null, $options6);
$worksheet3->setRow(9, RowColOptions::DEF_ROW_HEIGHT, null, $options6);
$worksheet3->setRow(10, RowColOptions::DEF_ROW_HEIGHT, null, $options7);
/* Write the sub-total data that is common to the row examples. */
$create_row_example_data($worksheet3, $bold);

/*
 * Example 4: Create a worksheet with outlined rows. Same as the example 1
 * except that the two sub-totals are collapsed.
 */
$options8 = new RowColOptions(1, 2, 0);
$options9 = new RowColOptions(1, 1, 1);
$options10 = new RowColOptions(0, 0, 1);
/* Set the row options with the outline level. */
$worksheet4->setRow(1, RowColOptions::DEF_ROW_HEIGHT, null, $options8);
$worksheet4->setRow(2, RowColOptions::DEF_ROW_HEIGHT, null, $options8);
$worksheet4->setRow(3, RowColOptions::DEF_ROW_HEIGHT, null, $options8);
$worksheet4->setRow(4, RowColOptions::DEF_ROW_HEIGHT, null, $options8);
$worksheet4->setRow(5, RowColOptions::DEF_ROW_HEIGHT, null, $options9);
$worksheet4->setRow(6, RowColOptions::DEF_ROW_HEIGHT, null, $options8);
$worksheet4->setRow(7, RowColOptions::DEF_ROW_HEIGHT, null, $options8);
$worksheet4->setRow(8, RowColOptions::DEF_ROW_HEIGHT, null, $options8);
$worksheet4->setRow(9, RowColOptions::DEF_ROW_HEIGHT, null, $options8);
$worksheet4->setRow(10, RowColOptions::DEF_ROW_HEIGHT, null, $options9);
$worksheet4->setRow(11, RowColOptions::DEF_ROW_HEIGHT, null, $options10);
/* Write the sub-total data that is common to the row examples. */
$create_row_example_data($worksheet4, $bold);

/*
 * Example 5: Create a worksheet with outlined columns.
 */
$options11 = new RowColOptions(0, 1, 0);
/* Write the sub-total data that is common to the column examples. */
$create_col_example_data($worksheet5, $bold);
/* Add bold format to the first row. */
$worksheet5->setRow(0, RowColOptions::DEF_ROW_HEIGHT, $bold);
/* Set column formatting and the outline level. */
$worksheet5->setColumn('A:A', 10, $bold);
$worksheet5->setColumn('B:G', 5, null, $options11);
$worksheet5->setColumn('H:H', 10);

/*
 * Example 6: Create a worksheet with outlined columns.
 */
$options12 = new RowColOptions(1, 1, 0);
$options13 = new RowColOptions(0, 0, 1);
/* Write the sub-total data that is common to the column examples. */
$create_col_example_data($worksheet6, $bold);
/* Add bold format to the first row. */
$worksheet6->setRow(0, RowColOptions::DEF_ROW_HEIGHT, $bold);
/* Set column formatting and the outline level. */
$worksheet6->setColumn('A:A', 10, $bold);
$worksheet6->setColumn('B:G', 5, null, $options12);
$worksheet6->setColumn('H:H', 10, null, $options13);

$workbook->close();
