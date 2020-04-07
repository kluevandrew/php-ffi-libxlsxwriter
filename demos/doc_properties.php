<?php

/**
 * @see http://libxlsxwriter.github.io/doc_properties_8c-example.html
 * @noinspection PhpUnhandledExceptionInspection
 */

use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Options\DocProperties;
use FFILibXlsxWriter\Workbook;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

$workbook = new Workbook(__DIR__ . '/output/doc_properties.xlsx');
$worksheet = $workbook->addWorksheet();

/* Create a properties structure and set some of the fields. */
$properties = new DocProperties();
$properties->title = 'This is an example spreadsheet';
$properties->subject = 'With document properties';
$properties->author = 'John McNamara';
$properties->manager = 'Dr. Heinz Doofenshmirtz';
$properties->company = 'of Wolves';
$properties->category = 'Example spreadsheets';
$properties->keywords = 'Sample, Example, Properties';
$properties->comments = 'Created with libxlsxwriter';
$properties->status = 'Quo';

$workbook->setProperties($properties);
/* Add some text to the file. */
$worksheet->setColumn('A:A', 50);
$worksheet->writeString('A!', 'Select \'Workbook Properties\' to see properties.');

$workbook->close();
