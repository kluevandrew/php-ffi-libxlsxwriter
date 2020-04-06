<?php

/**
 * @see http://libxlsxwriter.github.io/merge_rich_string_8c-example.html
 */

use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Enums\Align;
use FFILibXlsxWriter\Enums\Border;
use FFILibXlsxWriter\Enums\Color;
use FFILibXlsxWriter\Structs\RichString;
use FFILibXlsxWriter\Structs\RichStringPart;
use FFILibXlsxWriter\Workbook;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

$workbook = new Workbook(__DIR__ . '/output/merge_rich_string.xlsx');
$worksheet = $workbook->addWorksheet();

/* Configure a format for the merged range. */
$merge_format = $workbook->addFormat();
$merge_format->setAlign(Align::CENTER);
$merge_format->setAlign(Align::VERTICAL_CENTER);
$merge_format->setBorder(Border::THIN);
/* Configure formats for the rich string. */
$red = $workbook->addFormat();
$red->setFontColor(Color::RED);
$blue = $workbook->addFormat();
$blue->setFontColor(Color::BLUE);
/* Create the fragments for the rich string. */
$rich_string = new RichString([
    new RichStringPart('This is '),
    new RichStringPart('red', $red),
    new RichStringPart(' and this is '),
    new RichStringPart('blue', $blue),
]);

/* Write an empty string to the merged range. */
$worksheet->mergeRange(1, 1, 4, 3, "", $merge_format);
/* We then overwrite the first merged cell with a rich string. Note that
 * we must also pass the cell format used in the merged cells format at
 * the end. */
$worksheet->writeRichString(1, 1, $rich_string, $merge_format);

$workbook->close();
