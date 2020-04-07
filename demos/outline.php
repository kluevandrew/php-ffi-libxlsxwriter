<?php

/**
 * @see http://libxlsxwriter.github.io/outline_8c-example.html
 * @noinspection PhpUnhandledExceptionInspection
 * @noinspection DuplicatedCode
 */

use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Options\RowColOptions;
use FFILibXlsxWriter\Workbook;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

$workbook = new Workbook(__DIR__ . '/output/outline.xlsx');

$worksheet1 = $workbook->addWorksheet("Outlined Rows");
$worksheet2 = $workbook->addWorksheet("Collapsed Rows");
$worksheet3 = $workbook->addWorksheet("Outline Columns");
$worksheet4 = $workbook->addWorksheet("Outline levels");
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

/* Set the column width for clarity. */
$worksheet1->setColumn("A:A", 20);
/* Set the row options with the outline level. */
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
/* Add data and formulas to the worksheet. */
$worksheet1->writeString('A1', "Region", $bold);
$worksheet1->writeString('A2', "North");
$worksheet1->writeString('A3', "North");
$worksheet1->writeString('A4', "North");
$worksheet1->writeString('A5', "North");
$worksheet1->writeString('A6', "North Total", $bold);
$worksheet1->writeString('B1', "Sales", $bold);
$worksheet1->writeNumber('B2', 1000);
$worksheet1->writeNumber('B3', 1200);
$worksheet1->writeNumber('B4', 900);
$worksheet1->writeNumber('B5', 1200);
$worksheet1->writeFormula('B6', "=SUBTOTAL(9,B2:B5)", $bold);
$worksheet1->writeString('A7', "South");
$worksheet1->writeString('A8', "South");
$worksheet1->writeString('A9', "South");
$worksheet1->writeString('A10', "South");
$worksheet1->writeString('A11', "South Total", $bold);
$worksheet1->writeNumber('B7', 400);
$worksheet1->writeNumber('B8', 600);
$worksheet1->writeNumber('B9', 500);
$worksheet1->writeNumber('B10', 600);
$worksheet1->writeFormula('B11', "=SUBTOTAL(9,B7:B10)", $bold);
$worksheet1->writeString('A12', "Grand Total", $bold);
$worksheet1->writeFormula('B12', "=SUBTOTAL(9,B2:B10)", $bold);
/*
* Example 2: Create a worksheet with outlined rows. This is the same as
* the previous example except that the rows are collapsed.  Note: We need
* to indicate the row that contains the collapsed symbol '+' with the
* optional parameter, 'collapsed'.
*/
/* The option structs with the outline level and collapsed property set. */
$options3 = new RowColOptions(1, 2, 0);
$options4 = new RowColOptions(1, 1, 0);
$options5 = new RowColOptions(0, 0, 1);
/* Set the column width for clarity. */
$worksheet2->setColumn("A:A", 20);
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
/* Add data and formulas to the worksheet. */
$worksheet2->writeString('A1', "Region", $bold);
$worksheet2->writeString('A2', "North");
$worksheet2->writeString('A3', "North");
$worksheet2->writeString('A4', "North");
$worksheet2->writeString('A5', "North");
$worksheet2->writeString('A6', "North Total", $bold);
$worksheet2->writeString('B1', "Sales", $bold);
$worksheet2->writeNumber('B2', 1000);
$worksheet2->writeNumber('B3', 1200);
$worksheet2->writeNumber('B4', 900);
$worksheet2->writeNumber('B5', 1200);
$worksheet2->writeFormula('B6', "=SUBTOTAL(9,B2:B5)", $bold);
$worksheet2->writeString('A7', "South");
$worksheet2->writeString('A8', "South");
$worksheet2->writeString('A9', "South");
$worksheet2->writeString('A10', "South");
$worksheet2->writeString('A11', "South Total", $bold);
$worksheet2->writeNumber('B7', 400);
$worksheet2->writeNumber('B8', 600);
$worksheet2->writeNumber('B9', 500);
$worksheet2->writeNumber('B10', 600);
$worksheet2->writeFormula('B11', "=SUBTOTAL(9,B7:B10)", $bold);
$worksheet2->writeString('A12', "Grand Total", $bold);
$worksheet2->writeFormula('B12', "=SUBTOTAL(9,B2:B10)", $bold);
/*
 * Example 3: Create a worksheet with outlined columns.
 */
$options6 = new RowColOptions(0, 1, 0);
/* Add data and formulas to the worksheet. */
$worksheet3->writeString('A1', "Month");
$worksheet3->writeString('B1', "Jan");
$worksheet3->writeString('C1', "Feb");
$worksheet3->writeString('D1', "Mar");
$worksheet3->writeString('E1', "Apr");
$worksheet3->writeString('F1', "May");
$worksheet3->writeString('G1', "Jun");
$worksheet3->writeString('H1', "Total");
$worksheet3->writeString('A2', "North");
$worksheet3->writeNumber('B2', 50);
$worksheet3->writeNumber('C2', 20);
$worksheet3->writeNumber('D2', 15);
$worksheet3->writeNumber('E2', 25);
$worksheet3->writeNumber('F2', 65);
$worksheet3->writeNumber('G2', 80);
$worksheet3->writeFormula('H2', "=SUM(B2:G2)");
$worksheet3->writeString('A3', "South");
$worksheet3->writeNumber('B3', 10);
$worksheet3->writeNumber('C3', 20);
$worksheet3->writeNumber('D3', 30);
$worksheet3->writeNumber('E3', 50);
$worksheet3->writeNumber('F3', 50);
$worksheet3->writeNumber('G3', 50);
$worksheet3->writeFormula('H3', "=SUM(B3:G3)");
$worksheet3->writeString('A4', "East");
$worksheet3->writeNumber('B4', 45);
$worksheet3->writeNumber('C4', 75);
$worksheet3->writeNumber('D4', 50);
$worksheet3->writeNumber('E4', 15);
$worksheet3->writeNumber('F4', 75);
$worksheet3->writeNumber('G4', 100);
$worksheet3->writeFormula('H4', "=SUM(B4:G4)");
$worksheet3->writeString('A5', "West");
$worksheet3->writeNumber('B5', 15);
$worksheet3->writeNumber('C5', 15);
$worksheet3->writeNumber('D5', 55);
$worksheet3->writeNumber('E5', 35);
$worksheet3->writeNumber('F5', 20);
$worksheet3->writeNumber('G5', 50);
$worksheet3->writeFormula('H5', "=SUM(B5:G5)");
$worksheet3->writeFormula('H6', "=SUM(H2:H5)", $bold);
/* Add bold format to the first row. */
$worksheet3->setRow(0, RowColOptions::DEF_ROW_HEIGHT, $bold);
/* Set column formatting and the outline level. */
$worksheet3->setColumn("A:A", 10, $bold);
$worksheet3->setColumn("B:G", 5, null, $options6);
$worksheet3->setColumn("H:H", 10);
/*
 * Example 4: Show all possible outline levels.
 */
$level1 = new RowColOptions(0, 1, 0);
$level2 = new RowColOptions(0, 2, 0);
$level3 = new RowColOptions(0, 3, 0);
$level4 = new RowColOptions(0, 4, 0);
$level5 = new RowColOptions(0, 5, 0);
$level6 = new RowColOptions(0, 6, 0);
$level7 = new RowColOptions(0, 7, 0);
$worksheet4->writeString([0, 0], "Level 1");
$worksheet4->writeString([1, 0], "Level 2");
$worksheet4->writeString([2, 0], "Level 3");
$worksheet4->writeString([3, 0], "Level 4");
$worksheet4->writeString([4, 0], "Level 5");
$worksheet4->writeString([5, 0], "Level 6");
$worksheet4->writeString([6, 0], "Level 7");
$worksheet4->writeString([7, 0], "Level 6");
$worksheet4->writeString([8, 0], "Level 5");
$worksheet4->writeString([9, 0], "Level 4");
$worksheet4->writeString([10, 0], "Level 3");
$worksheet4->writeString([11, 0], "Level 2");
$worksheet4->writeString([12, 0], "Level 1");
$worksheet4->setRow(0, RowColOptions::DEF_ROW_HEIGHT, null, $level1);
$worksheet4->setRow(1, RowColOptions::DEF_ROW_HEIGHT, null, $level2);
$worksheet4->setRow(2, RowColOptions::DEF_ROW_HEIGHT, null, $level3);
$worksheet4->setRow(3, RowColOptions::DEF_ROW_HEIGHT, null, $level4);
$worksheet4->setRow(4, RowColOptions::DEF_ROW_HEIGHT, null, $level5);
$worksheet4->setRow(5, RowColOptions::DEF_ROW_HEIGHT, null, $level6);
$worksheet4->setRow(6, RowColOptions::DEF_ROW_HEIGHT, null, $level7);
$worksheet4->setRow(7, RowColOptions::DEF_ROW_HEIGHT, null, $level6);
$worksheet4->setRow(8, RowColOptions::DEF_ROW_HEIGHT, null, $level5);
$worksheet4->setRow(9, RowColOptions::DEF_ROW_HEIGHT, null, $level4);
$worksheet4->setRow(10, RowColOptions::DEF_ROW_HEIGHT, null, $level3);
$worksheet4->setRow(11, RowColOptions::DEF_ROW_HEIGHT, null, $level2);
$worksheet4->setRow(12, RowColOptions::DEF_ROW_HEIGHT, null, $level1);

$workbook->close();
