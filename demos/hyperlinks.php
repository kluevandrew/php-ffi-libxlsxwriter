<?php

/**
 * @see http://libxlsxwriter.github.io/hyperlinks_8c-example.html
 * @noinspection PhpUnhandledExceptionInspection
 */

use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Enums\Color;
use FFILibXlsxWriter\Enums\Underline;
use FFILibXlsxWriter\Workbook;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

$workbook = new Workbook(__DIR__ . '/output/hyperlinks.xlsx');
$worksheet = $workbook->addWorksheet();

/* Get the default url format (used in the overwriting examples below). */
$urlFormat = $workbook->getDefaultUrlFormat();
/* Create a user defined link format. */
$redFormat = $workbook->addFormat();
$redFormat->setUnderline(Underline::SINGLE);
$redFormat->setFontColor(Color::RED);

/* Widen the first column to make the text clearer. */
$worksheet->setColumn([0, 0], 30, null);

/* Write a hyperlink. A default blue underline will be used if the format is null. */
$worksheet->writeUrl([0, 0], 'https://github.com/kluevandrew/php-ffi-libxlsxwriter', null);

/* Write a hyperlink but overwrite the displayed string. Note, we need to
 * specify the format for the string to match the default hyperlink. */
$worksheet->writeUrl([2, 0], "https://github.com/kluevandrew/php-ffi-libxlsxwriter", null);
$worksheet->writeString([2, 0], "Read the documentation.", $urlFormat);

/* Write a hyperlink with a different format. */
$worksheet->writeUrl([4, 0], "https://github.com/kluevandrew/php-ffi-libxlsxwriter", $redFormat);

/* Write a mail hyperlink. */
$worksheet->writeUrl([6, 0], "mailto:kluev.andrew@gmail.com", null);

/* Write a mail hyperlink and overwrite the displayed string. We again
 * specify the format for the string to match the default hyperlink. */
$worksheet->writeUrl([8, 0], "mailto:kluev.andrew@gmail.com", null);
$worksheet->writeString([8, 0], "Drop me a line.", $urlFormat);

$workbook->close();
