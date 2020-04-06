<?php

/**
 * @noinspection PhpUnhandledExceptionInspection
 */

use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Enums\Align;
use FFILibXlsxWriter\Enums\Color;
use FFILibXlsxWriter\Enums\Font;
use FFILibXlsxWriter\Workbook;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

$workbook = new Workbook(__DIR__ . '/output/styling.xlsx');
$worksheet = $workbook->addWorksheet('First worksheet');

$bold = $workbook->addFormat();
$bold->setBold();

$italic = $workbook->addFormat();
$italic->setItalic();

$red = $workbook->addFormat();
$red->setFontColor(Color::RED);

$centered = $workbook->addFormat();
$centered->setAlign(Align::CENTER);

$superscript = $workbook->addFormat();
$superscript->setFontScript(Font::SUPERSCRIPT);

$worksheet->writeString([0, 0], 'This is bold', $bold);
$worksheet->writeString([1, 0], 'This is italic', $bold);
$worksheet->writeString([2, 0], 'This is red', $red);
$worksheet->writeString([3, 0], 'This is centered', $centered);
$worksheet->writeString([4, 0], 'This is superscript', $superscript);

$worksheet->setColumn('A:A', 30);

$workbook->close();
