<?php

/**
 * @see http://libxlsxwriter.github.io/defined_name_8c-example.html
 * @noinspection PhpUnhandledExceptionInspection
 */

use FFILibXlsxWriter\Exceptions\FFILibXlsxWriterException;
use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Workbook;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

$workbook = new Workbook(__DIR__ . '/output/defined_name.xlsx');

/* We don't use the returned worksheets in this example and use a generic
 * loop instead. */
$workbook->addWorksheet();
$workbook->addWorksheet();

/* Define some global/workbook names. */
$workbook->defineName('Sales', '=!G1:H10');
$workbook->defineName('Exchange_rate', '=0.96');
try {
    $workbook->defineName('Sales', '=Sheet1!$G$1:$H$10');
} catch (FFILibXlsxWriterException $exception) {
    // already defined
}

/* Define a local/worksheet name. */
$workbook->defineName('Sheet2!Sales', '=Sheet2!$G$1:$G$10');

/* Write some text to the worksheets and one of the defined name in a formula. */
foreach ($workbook->getWorksheets() as $worksheet) {
    $worksheet->setColumn('A:A', 45);
    $worksheet->writeString([0, 0], 'This worksheet contains some defined names.');
    $worksheet->writeString([1, 0], 'See Formulas -> Name Manager above.');
    $worksheet->writeString([2, 0], 'Example formula in cell B3 ->');
    $worksheet->writeFormula([2, 1], '=Exchange_rate');
}
$workbook->close();
