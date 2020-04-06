<?php

/**
 * @see http://libxlsxwriter.github.io/headers_footers_8c-example.html
 * @noinspection PhpUnhandledExceptionInspection
 */

use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Workbook;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

$workbook = new Workbook(__DIR__ . '/output/headers_footers.xlsx');
$preview = "Select Print Preview to see the header and footer";

/*
 * A simple example to start
 */
$worksheet1 = $workbook->addWorksheet("Simple");
$header1 = "&CHere is some centered text.";
$footer1 = "&LHere is some left aligned text.";
$worksheet1->setHeader($header1);
$worksheet1->setFooter($footer1);
$worksheet1->setColumn([0, 0], 50);
$worksheet1->writeString([0, 0], $preview);

/*
 * This is an example of some of the header/footer variables.
 */
$worksheet2 = $workbook->addWorksheet("Variables");
$header2 = "&LPage &P of &N &CFilename: &F &RSheetname: &A";
$footer2 = "&LCurrent date: &D &RCurrent time: &T";
$breaks = [20, 0];
$worksheet2->setHeader($header2);
$worksheet2->setFooter($footer2);
$worksheet2->setColumn([0, 0], 50);
$worksheet2->writeString([0, 0], $preview);
$worksheet2->setHorizontalBreaks($breaks);
$worksheet2->writeString([20, 0], "Next page");

/*
 * This example shows how to use more than one font.
 */
$worksheet3 = $workbook->addWorksheet("Mixed fonts");
$header3 = "&C&\"Courier New,Bold\"Hello &\"Arial,Italic\"World";
$footer3 = "&C&\"Symbol\"e&\"Arial\" = mc&X2";
$worksheet3->setHeader($header3);
$worksheet3->setFooter($footer3);
$worksheet3->setColumn([0, 0], 50);
$worksheet3->writeString([0, 0], $preview);

/*
 * Example of line wrapping.
 */
$worksheet4 = $workbook->addWorksheet("Word wrap");
$header4 = "&CHeading 1\nHeading 2";
$worksheet4->setHeader($header4);
$worksheet4->setColumn([0, 0], 50);
$worksheet4->writeString([0, 0], $preview);

/*
 * Example of inserting a literal ampersand &
 */
$worksheet5 = $workbook->addWorksheet("Ampersand");
$header5 = "&CCuriouser && Curiouser - Attorneys at Law";
$worksheet5->setHeader($header5);
$worksheet5->setColumn([0, 0], 50);
$worksheet5->writeString([0, 0], $preview);

$workbook->close();
