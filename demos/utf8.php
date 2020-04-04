<?php
/**
 * @see http://libxlsxwriter.github.io/utf8_8c-example.html
 */
use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Workbook;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

$workbook = new Workbook(__DIR__ . '/output/utf8.xlsx');
$worksheet = $workbook->addWorksheet();

$worksheet->writeString( 2, 1, 'Это фраза на русском!', null);

$workbook->close();
