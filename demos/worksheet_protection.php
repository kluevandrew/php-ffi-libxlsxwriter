<?php
/**
 * @see http://libxlsxwriter.github.io/worksheet_protection_8c-example.html
 */

use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Workbook;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

$workbook = new Workbook(__DIR__ . '/output/worksheet_protection.xlsx');
$worksheet = $workbook->addWorksheet();

$unlocked = $workbook->addFormat();
$unlocked->setUnlocked();
$hidden = $workbook->addFormat();
$hidden->setHidden();
/* Widen the first column to make the text clearer. */
$worksheet->setColumn(0, 0, 40, NULL);
/* Turn worksheet protection on without a password. */
$worksheet->protect($password = null, $options = null);
/* Write a locked, unlocked and hidden cell. */
$worksheet->writeString(0, 0, "B1 is locked. It cannot be edited.");
$worksheet->writeString(1, 0, "B2 is unlocked. It can be edited.");
$worksheet->writeString(2, 0, "B3 is hidden. The formula isn't visible.");
$worksheet->writeFormula(0, 1, "=1+2"); /* Locked by default. */
$worksheet->writeFormula(1, 1, "=1+2", $unlocked);
$worksheet->writeFormula(2, 1, "=1+2", $hidden);

$workbook->close();
