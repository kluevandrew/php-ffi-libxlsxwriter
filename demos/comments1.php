<?php

/**
 * @see http://libxlsxwriter.github.io/comments1_8c-example.html
 * @noinspection PhpUnhandledExceptionInspection
 */

use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Workbook;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

$workbook = new Workbook(__DIR__ . '/output/comments1.xlsx');
$worksheet = $workbook->addWorksheet();

$worksheet->writeString('A1', 'Hello');
$worksheet->writeComment('A1', 'This is a comment');

$workbook->close();
