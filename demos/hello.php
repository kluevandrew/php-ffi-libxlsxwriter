<?php

/**
 * @see http://libxlsxwriter.github.io/hello_8c-example.html
 * @noinspection PhpUnhandledExceptionInspection
 */

use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Workbook;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

$workbook = new Workbook(__DIR__ . '/output/hello.xlsx');
$worksheet = $workbook->addWorksheet();

$worksheet->writeString([0, 0], "Hello");
$worksheet->writeNumber([1, 0], 123);

$workbook->close();
