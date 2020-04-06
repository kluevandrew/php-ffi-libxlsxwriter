<?php

/**
 * @see http://libxlsxwriter.github.io/rich_strings_8c-example.html
 */

use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Enums\Align;
use FFILibXlsxWriter\Enums\Color;
use FFILibXlsxWriter\Enums\Font;
use FFILibXlsxWriter\Structs\RichString;
use FFILibXlsxWriter\Structs\RichStringPart;
use FFILibXlsxWriter\Workbook;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

$workbook = new Workbook(__DIR__ . '/output/rich_strings.xlsx');
$worksheet = $workbook->addWorksheet();

/* Set up some formats to use. */
$bold = $workbook->addFormat();
$bold->setBold();
$italic = $workbook->addFormat();
$italic->setItalic();
$red = $workbook->addFormat();
$red->setFontColor(Color::RED);
$blue = $workbook->addFormat();
$blue->setFontColor(Color::BLUE);
$center = $workbook->addFormat();
$center->setAlign(Align::CENTER);
$superscript = $workbook->addFormat();
$superscript->setFontScript(Font::SUPERSCRIPT);
/* Make the first column wider for clarity. */
$worksheet->setColumn(0, 0, 30);
/*
 * Create and write some rich strings with multiple formats.
 */
/* Example 1. Some bold and italic in the same string. */
$rich_string1 = new RichString([
    new RichStringPart('This is '),
    new RichStringPart('bold', $bold),
    new RichStringPart(' and this is '),
    new RichStringPart('italic', $italic),
]);
list($row, $col) = FFILibXlsxWriter::cell('A1');
$worksheet->writeRichString($row, $col, $rich_string1, null);

/* Example 2. Some red and blue coloring in the same string. */
$rich_string2 = new RichString([
    new RichStringPart('This is '),
    new RichStringPart('red', $red),
    new RichStringPart(' and this is '),
    new RichStringPart('blue', $blue),
]);
list($row, $col) = FFILibXlsxWriter::cell('A3');
$worksheet->writeRichString($row, $col, $rich_string2, null);

/* Example 3. A rich string plus cell formatting. */
$rich_string3 = new RichString([
    new RichStringPart('Some '),
    new RichStringPart('bold text', $bold),
    new RichStringPart(' centered'),
]);
list($row, $col) = FFILibXlsxWriter::cell('A5');
$worksheet->writeRichString($row, $col, $rich_string3, $center);

/* Example 4. A math example with a superscript. */
$rich_string4 = new RichString([
    new RichStringPart('j =k', $italic),
    new RichStringPart('(n-1)', $superscript),
]);
list($row, $col) = FFILibXlsxWriter::cell('A7');
$worksheet->writeRichString($row, $col, $rich_string4, $center);

$workbook->close();
