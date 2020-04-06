<?php

/**
 * @see http://libxlsxwriter.github.io/data_validate_8c-example.html
 */

use FFILibXlsxWriter\Enums\ValidationCriteria;
use FFILibXlsxWriter\Enums\ValidationErrorType;
use FFILibXlsxWriter\Enums\ValidationType;
use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Enums\Align;
use FFILibXlsxWriter\Enums\Border;
use FFILibXlsxWriter\Structs\DataValidation;
use FFILibXlsxWriter\Style\Format;
use FFILibXlsxWriter\Workbook;
use FFILibXlsxWriter\Worksheet;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

/*
 * Write some data to the worksheet.
 */
$write_worksheet_data = function (Worksheet $worksheet, Format $format): void {
    list($row, $col) = FFILibXlsxWriter::cell("A1");
    $worksheet->writeString($row, $col, "Some examples of data validation in libxlsxwriter", $format);
    list($row, $col) = FFILibXlsxWriter::cell("B1");
    $worksheet->writeString($row, $col, "Enter values in this column", $format);
    list($row, $col) = FFILibXlsxWriter::cell("D1");
    $worksheet->writeString($row, $col, "Sample Data", $format);
    list($row, $col) = FFILibXlsxWriter::cell("D3");
    $worksheet->writeString($row, $col, "Integers");
    list($row, $col) = FFILibXlsxWriter::cell("E3");
    $worksheet->writeNumber($row, $col, 1);
    list($row, $col) = FFILibXlsxWriter::cell("F3");
    $worksheet->writeNumber($row, $col, 10);
    list($row, $col) = FFILibXlsxWriter::cell("D4");
    $worksheet->writeString($row, $col, "List Data");
    list($row, $col) = FFILibXlsxWriter::cell("E4");
    $worksheet->writeString($row, $col, "open");
    list($row, $col) = FFILibXlsxWriter::cell("F4");
    $worksheet->writeString($row, $col, "high");
    list($row, $col) = FFILibXlsxWriter::cell("G4");
    $worksheet->writeString($row, $col, "close");
    list($row, $col) = FFILibXlsxWriter::cell("D5");
    $worksheet->writeString($row, $col, "Formula");
    list($row, $col) = FFILibXlsxWriter::cell("E5");
    $worksheet->writeFormula($row, $col, "=AND(F5=50,G5=60)");
    list($row, $col) = FFILibXlsxWriter::cell("F5");
    $worksheet->writeNumber($row, $col, 50);
    list($row, $col) = FFILibXlsxWriter::cell("G5");
    $worksheet->writeNumber($row, $col, 60);
};

$workbook = new Workbook(__DIR__ . '/output/data_validate.xlsx');
$worksheet = $workbook->addWorksheet();
$data_validation = new DataValidation();

/* Add a format to use to highlight the header cells. */
$format = $workbook->addFormat();
$format->setBorder(Border::THIN);
$format->setFgColor(0xC6EFCE);
$format->setBold();
$format->setTextWrap();
$format->setAlign(Align::VERTICAL_CENTER);
$format->setIndent(1);
/* Write some data for the validations. */
$write_worksheet_data($worksheet, $format);
/* Set up layout of the worksheet. */
$worksheet->setColumn(0, 0, 55);
$worksheet->setColumn(1, 1, 15);
$worksheet->setColumn(3, 3, 15);
$worksheet->setRow(0, 36);

/*
* Example 1. Limiting input to an integer in a fixed range.
*/
list($row, $col) = FFILibXlsxWriter::cell('A3');
$worksheet->writeString($row, $col, "Enter an integer between 1 and 10");
$data_validation->validate = ValidationType::INTEGER;
$data_validation->criteria = ValidationCriteria::BETWEEN;
$data_validation->minimum_number = 1;
$data_validation->maximum_number = 10;
list($row, $col) = FFILibXlsxWriter::cell('B3');
$worksheet->dataValidationCell($row, $col, $data_validation);

/*
 * Example 2. Limiting input to an integer outside a fixed range.
 */
$data_validation = new DataValidation();
list($row, $col) = FFILibXlsxWriter::cell('A5');
$worksheet->writeString($row, $col, "Enter an integer not between 1 and 10 (using cell references)");
//$data_validation = new DataValidation();
$data_validation->validate = ValidationType::INTEGER_FORMULA;
$data_validation->criteria = ValidationCriteria::BETWEEN;
$data_validation->minimum_formula = '=E3';
$data_validation->maximum_formula = '=F3';

list($row, $col) = FFILibXlsxWriter::cell('B5');
$worksheet->dataValidationCell($row, $col, $data_validation);

/*
 * Example 3. Limiting input to an integer greater than a fixed value.
 */
$data_validation = new DataValidation();
list($row, $col) = FFILibXlsxWriter::cell('A7');
$worksheet->writeString($row, $col, "Enter an integer greater than 0");
$data_validation->validate = ValidationType::INTEGER;
$data_validation->criteria = ValidationCriteria::GREATER_THAN;
list($row, $col) = FFILibXlsxWriter::cell('B7');
$worksheet->dataValidationCell($row, $col, $data_validation);

/*
 * Example 4. Limiting input to an integer less than a fixed value.
 */
$data_validation = new DataValidation();
list($row, $col) = FFILibXlsxWriter::cell('A9');
$worksheet->writeString($row, $col, "Enter an integer less than 10");
$data_validation->validate = ValidationType::INTEGER;
$data_validation->criteria = ValidationCriteria::LESS_THAN;
$data_validation->value_number = 10;
list($row, $col) = FFILibXlsxWriter::cell('B9');
$worksheet->dataValidationCell($row, $col, $data_validation);

/*
 * Example 5. Limiting input to a decimal in a fixed range.
 */
$data_validation = new DataValidation();
list($row, $col) = FFILibXlsxWriter::cell('A11');
$worksheet->writeString($row, $col, "Enter a decimal between 0.1 and 0.5");
$data_validation->validate = ValidationType::DECIMAL;
$data_validation->criteria = ValidationCriteria::BETWEEN;
$data_validation->minimum_number = 0.1;
$data_validation->maximum_number = 0.5;
list($row, $col) = FFILibXlsxWriter::cell('B11');
$worksheet->dataValidationCell($row, $col, $data_validation);

/*
 * Example 6. Limiting input to a value in a dropdown list.
 */
$data_validation = new DataValidation();
list($row, $col) = FFILibXlsxWriter::cell('A13');
$worksheet->writeString($row, $col, 'Select a value from a drop down list');
$data_validation->validate = ValidationType::LIST;
$data_validation->value_list = ["open", "high", "close"];
list($row, $col) = FFILibXlsxWriter::cell('B13');
$worksheet->dataValidationCell($row, $col, $data_validation);

/*
 * Example 7. Limiting input to a value in a dropdown list.
 */
$data_validation = new DataValidation();
list($row, $col) = FFILibXlsxWriter::cell('A15');
$worksheet->writeString($row, $col, 'Select a value from a drop down list (using a cell range)');
$data_validation->validate = ValidationType::LIST_FORMULA;
$data_validation->value_formula = '=$E$4:$G$4';
list($row, $col) = FFILibXlsxWriter::cell('B15');
$worksheet->dataValidationCell($row, $col, $data_validation);

/*
* Example 8. Limiting input to a date in a fixed range.
*/
$data_validation = new DataValidation();
list($row, $col) = FFILibXlsxWriter::cell('A17');
$worksheet->writeString($row, $col, "Enter a date between 1/1/2008 and 12/12/2008");
$datetime1 = new \FFILibXlsxWriter\Structs\DateTime(2008, 1, 1, 0, 0, 0);
$datetime2 = new \FFILibXlsxWriter\Structs\DateTime(2008, 12, 12, 0, 0, 0);
$data_validation->validate = ValidationType::DATE;
$data_validation->criteria = ValidationCriteria::BETWEEN;
$data_validation->minimum_datetime = $datetime1;
$data_validation->maximum_datetime = $datetime2;
list($row, $col) = FFILibXlsxWriter::cell('B17');
$worksheet->dataValidationCell($row, $col, $data_validation);

/*
 * Example 9. Limiting input to a time in a fixed range.
 */
$data_validation = new DataValidation();
list($row, $col) = FFILibXlsxWriter::cell('A19');
$worksheet->writeString($row, $col, "Enter a time between 6:00 and 12:00");
$datetime3 = new \FFILibXlsxWriter\Structs\DateTime(0, 0, 0, 6, 0, 0);
$datetime4 = new \FFILibXlsxWriter\Structs\DateTime(0, 0, 0, 12, 0, 0);
$data_validation->validate = ValidationType::DATE;
$data_validation->criteria = ValidationCriteria::BETWEEN;
$data_validation->minimum_datetime = $datetime3;
$data_validation->maximum_datetime = $datetime4;
list($row, $col) = FFILibXlsxWriter::cell('B19');
$worksheet->dataValidationCell($row, $col, $data_validation);

/*
 * Example 10. Limiting input to a string greater than a fixed length.
 */
$data_validation = new DataValidation();
list($row, $col) = FFILibXlsxWriter::cell('A21');
$worksheet->writeString($row, $col, "Enter a string longer than 3 characters");
$data_validation->validate = ValidationType::LENGTH;
$data_validation->criteria = ValidationCriteria::GREATER_THAN;
$data_validation->value_number = 3;
list($row, $col) = FFILibXlsxWriter::cell('B21');
$worksheet->dataValidationCell($row, $col, $data_validation);

/*
 * Example 11. Limiting input based on a formula.
 */
$data_validation = new DataValidation();
list($row, $col) = FFILibXlsxWriter::cell('A23');
$worksheet->writeString($row, $col, 'Enter a value if the following is true "=AND(F5=50,G5=60)"');
$data_validation->validate = ValidationType::CUSTOM_FORMULA;
$data_validation->value_formula = "=AND(F5=50,G5=60)";
list($row, $col) = FFILibXlsxWriter::cell('B23');
$worksheet->dataValidationCell($row, $col, $data_validation);

/*
 * Example 12. Displaying and modifying data validation messages.
 */
$data_validation = new DataValidation();
list($row, $col) = FFILibXlsxWriter::cell('A25');
$worksheet->writeString($row, $col, "Displays a message when you select the cell");
$data_validation->validate = ValidationType::INTEGER;
$data_validation->criteria = ValidationCriteria::BETWEEN;
$data_validation->minimum_number = 1;
$data_validation->maximum_number = 100;
$data_validation->input_title = "Enter an integer:";
$data_validation->input_message = "between 1 and 100";
list($row, $col) = FFILibXlsxWriter::cell('B25');
$worksheet->dataValidationCell($row, $col, $data_validation);

/*
 * Example 13. Displaying and modifying data validation messages.
 */
$data_validation = new DataValidation();
list($row, $col) = FFILibXlsxWriter::cell('A27');
$worksheet->writeString($row, $col, "Display a custom error message when integer isn't between 1 and 100");
$data_validation->validate = ValidationType::INTEGER;
$data_validation->criteria = ValidationCriteria::BETWEEN;
$data_validation->minimum_number = 1;
$data_validation->maximum_number = 100;
$data_validation->input_title = "Enter an integer:";
$data_validation->input_message = "between 1 and 100";
$data_validation->error_title = "Input value is not valid!";
$data_validation->error_message = "It should be an integer between 1 and 100";
list($row, $col) = FFILibXlsxWriter::cell('B27');
$worksheet->dataValidationCell($row, $col, $data_validation);

/*
 * Example 14. Displaying and modifying data validation messages.
 */
$data_validation = new DataValidation();
list($row, $col) = FFILibXlsxWriter::cell('A29');
$worksheet->writeString($row, $col, "Display a custom info message when integer isn't between 1 and 100");
$data_validation->validate = ValidationType::INTEGER;
$data_validation->criteria = ValidationCriteria::BETWEEN;
$data_validation->minimum_number = 1;
$data_validation->maximum_number = 100;
$data_validation->input_title = "Enter an integer:";
$data_validation->input_message = "between 1 and 100";
$data_validation->error_title = "Input value is not valid!";
$data_validation->error_message = "It should be an integer between 1 and 100";
$data_validation->error_type = ValidationErrorType::INFORMATION;
list($row, $col) = FFILibXlsxWriter::cell('B29');
$worksheet->dataValidationCell($row, $col, $data_validation);

/* Cleanup. */
unset($data_validation);
$workbook->close();
