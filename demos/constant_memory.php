<?php
/**
 * @see http://libxlsxwriter.github.io/constant_memory_8c-example.html
 */
use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Structs\WorkbookOptions;
use FFILibXlsxWriter\Workbook;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

$row = 0;
$col = 0;
$maxRow = 1000;
$maxCol = 50;
/* Set the worksheet options. */
$options = new WorkbookOptions(
    $constantMemory = true,
    $useZip64 = false,
    $tmpdir = __DIR__ . '/output',
);
/* Create a new workbook with options. */
$workbook = new Workbook(__DIR__ . '/output/constant_memory.xlsx', $options);
$worksheet = $workbook->addWorksheet();


for ($row = 0; $row < $maxRow; $row++) {
    for ($col = 0; $col < $maxCol; $col++) {
        $worksheet->writeNumber($row, $col, 123.45, null);
    }
}

$workbook->close();
