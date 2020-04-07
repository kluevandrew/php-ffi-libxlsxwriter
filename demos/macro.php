<?php

/**
 * @see http://libxlsxwriter.github.io/macro_8c-example.html
 * @noinspection PhpUnhandledExceptionInspection
 */

use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Workbook;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

$workbook = new Workbook(__DIR__ . '/output/macro.xlsm');
$worksheet = $workbook->addWorksheet();

/* Add a macro that will execute when the file is opened. */
$workbook->addVbaProject(__DIR__ . '/vbaProject.bin');
$workbook->setVbaName('my_name_is_amazing');
$worksheet->writeString('A1', 'You can find macros in menu');

$workbook->close();
