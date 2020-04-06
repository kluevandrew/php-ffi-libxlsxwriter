<?php

/**
 * @see http://libxlsxwriter.github.io/data_validate_8c-example.html
 * @noinspection PhpUnhandledExceptionInspection
 * @noinspection DuplicatedCode
 */

use FFILibXlsxWriter\Enums\ValidationCriteria;
use FFILibXlsxWriter\Enums\ValidationErrorType;
use FFILibXlsxWriter\Enums\ValidationType;
use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Enums\Align;
use FFILibXlsxWriter\Enums\Border;
use FFILibXlsxWriter\Structs\DataValidation;
use FFILibXlsxWriter\Structs\DateTime as XlsxDateTime;
use FFILibXlsxWriter\Style\Format;
use FFILibXlsxWriter\Workbook;
use FFILibXlsxWriter\Worksheet;

require_once __DIR__ . '/../vendor/autoload.php';

FFILibXlsxWriter::init();

/*
 * Write some data to the worksheet.
 */
$write_worksheet_data = function (Worksheet $worksheet, Format $format): void {
    $worksheet->writeString('A1', 'Some examples of data validation in libxlsxwriter', $format);
    $worksheet->writeString('B1', 'Enter values in this column', $format);
    $worksheet->writeString('D1', 'Sample Data', $format);
    $worksheet->writeString('D3', 'Integers');
    $worksheet->writeNumber('E3', 1);
    $worksheet->writeNumber('F3', 10);
    $worksheet->writeString('D4', 'List Data');
    $worksheet->writeString('E4', 'open');
    $worksheet->writeString('F4', 'high');
    $worksheet->writeString('G4', 'close');
    $worksheet->writeString('D5', 'Formula');
    $worksheet->writeFormula('E5', '=AND(F5=50,G5=60)');
    $worksheet->writeNumber('F5', 50);
    $worksheet->writeNumber('G5', 60);
};

$workbook = new Workbook(__DIR__ . '/output/data_validate.xlsx');
$worksheet = $workbook->addWorksheet();

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
$worksheet->setColumn([0, 0], 55);
$worksheet->setColumn([1, 1], 15);
$worksheet->setColumn([3, 3], 15);
$worksheet->setRow(0, 36);

/*
* Example 1. Limiting input to an integer in a fixed range.
*/
$worksheet->writeString('A3', 'Enter an integer between 1 and 10');
$data_validation = new DataValidation(ValidationCriteria::BETWEEN);
$data_validation->validate = ValidationType::INTEGER;
$data_validation->minimum_number = 1;
$data_validation->maximum_number = 10;
$worksheet->dataValidationCell('B3', $data_validation);

/*
 * Example 2. Limiting input to an integer outside a fixed range.
 */
$worksheet->writeString('A5', 'Enter an integer not between 1 and 10 (using cell references)');
$data_validation = new DataValidation(ValidationCriteria::BETWEEN);
$data_validation->validate = ValidationType::INTEGER_FORMULA;
$data_validation->minimum_formula = '=E3';
$data_validation->maximum_formula = '=F3';
$worksheet->dataValidationCell('B5', $data_validation);

/*
 * Example 3. Limiting input to an integer greater than a fixed value.
 */
$worksheet->writeString('A7', 'Enter an integer greater than 0');
$data_validation = new DataValidation(ValidationCriteria::GREATER_THAN);
$data_validation->validate = ValidationType::INTEGER;
$worksheet->dataValidationCell('B7', $data_validation);

/*
 * Example 4. Limiting input to an integer less than a fixed value.
 */
$worksheet->writeString('A9', 'Enter an integer less than 10');
$data_validation = new DataValidation(ValidationCriteria::LESS_THAN);
$data_validation->validate = ValidationType::INTEGER;
$data_validation->value_number = 10;
$worksheet->dataValidationCell('B9', $data_validation);

/*
 * Example 5. Limiting input to a decimal in a fixed range.
 */
$worksheet->writeString('A11', 'Enter a decimal between 0.1 and 0.5');
$data_validation = new DataValidation(ValidationCriteria::BETWEEN);
$data_validation->validate = ValidationType::DECIMAL;
$data_validation->minimum_number = 0.1;
$data_validation->maximum_number = 0.5;
$worksheet->dataValidationCell('B11', $data_validation);

/*
 * Example 6. Limiting input to a value in a dropdown list.
 */
$worksheet->writeString('A13', /** @lang text */ 'Select a value from a drop down list');
$data_validation = new DataValidation(ValidationCriteria::NONE);
$data_validation->validate = ValidationType::LIST;
$data_validation->value_list = ['open', 'high', 'close'];
$worksheet->dataValidationCell('B13', $data_validation);

/*
 * Example 7. Limiting input to a value in a dropdown list.
 */
$worksheet->writeString('A15', /** @lang text */ 'Select a value from a drop down list (using a cell range)');
$data_validation = new DataValidation(ValidationCriteria::NONE);
$data_validation->validate = ValidationType::LIST_FORMULA;
$data_validation->value_formula = '=$E$4:$G$4';
$worksheet->dataValidationCell('B15', $data_validation);

/*
* Example 8. Limiting input to a date in a fixed range.
*/
$worksheet->writeString('A17', 'Enter a date between 1/1/2008 and 12/12/2008');
$data_validation = new DataValidation(ValidationCriteria::BETWEEN);
$data_validation->validate = ValidationType::DATE;
$data_validation->minimum_datetime = new XlsxDateTime(2008, 1, 1, 0, 0, 0);
$data_validation->maximum_datetime = new XlsxDateTime(2008, 12, 12, 0, 0, 0);
$worksheet->dataValidationCell('B17', $data_validation);

/*
 * Example 9. Limiting input to a time in a fixed range.
 */
$worksheet->writeString('A19', 'Enter a time between 6:00 and 12:00');
$data_validation = new DataValidation(ValidationCriteria::BETWEEN);
$data_validation->validate = ValidationType::DATE;
$data_validation->minimum_datetime = new XlsxDateTime(0, 0, 0, 6, 0, 0);
$data_validation->maximum_datetime = new XlsxDateTime(0, 0, 0, 12, 0, 0);
$worksheet->dataValidationCell('B19', $data_validation);

/*
 * Example 10. Limiting input to a string greater than a fixed length.
 */
$worksheet->writeString('A21', 'Enter a string longer than 3 characters');
$data_validation = new DataValidation(ValidationCriteria::GREATER_THAN);
$data_validation->validate = ValidationType::LENGTH;
$data_validation->value_number = 3;
$worksheet->dataValidationCell('B21', $data_validation);

/*
 * Example 11. Limiting input based on a formula.
 */
$worksheet->writeString('A23', 'Enter a value if the following is true "=AND(F5=50,G5=60)"');
$data_validation = new DataValidation();
$data_validation->validate = ValidationType::CUSTOM_FORMULA;
$data_validation->value_formula = '=AND(F5=50,G5=60)';
$worksheet->dataValidationCell('B23', $data_validation);

/*
 * Example 12. Displaying and modifying data validation messages.
 */
$worksheet->writeString('A25', 'Displays a message when you select the cell');
$data_validation = new DataValidation(ValidationCriteria::BETWEEN);
$data_validation->validate = ValidationType::INTEGER;
$data_validation->minimum_number = 1;
$data_validation->maximum_number = 100;
$data_validation->input_title = 'Enter an integer:';
$data_validation->input_message = 'between 1 and 100';
$worksheet->dataValidationCell('B25', $data_validation);

/*
 * Example 13. Displaying and modifying data validation messages.
 */
$worksheet->writeString('A27', 'Display a custom error message when integer isn\'t between 1 and 100');
$data_validation = new DataValidation(ValidationCriteria::BETWEEN);
$data_validation->validate = ValidationType::INTEGER;
$data_validation->minimum_number = 1;
$data_validation->maximum_number = 100;
$data_validation->input_title = 'Enter an integer:';
$data_validation->input_message = 'between 1 and 100';
$data_validation->error_title = 'Input value is not valid!';
$data_validation->error_message = 'It should be an integer between 1 and 100';
$worksheet->dataValidationCell('B27', $data_validation);

/*
 * Example 14. Displaying and modifying data validation messages.
 */
$worksheet->writeString('A29', 'Display a custom info message when integer isn\'t between 1 and 100');
$data_validation = new DataValidation(ValidationCriteria::BETWEEN);
$data_validation->validate = ValidationType::INTEGER;
$data_validation->minimum_number = 1;
$data_validation->maximum_number = 100;
$data_validation->input_title = 'Enter an integer:';
$data_validation->input_message = 'between 1 and 100';
$data_validation->error_title = 'Input value is not valid!';
$data_validation->error_message = 'It should be an integer between 1 and 100';
$data_validation->error_type = ValidationErrorType::INFORMATION;
$worksheet->dataValidationCell('B29', $data_validation);

/* Cleanup. */
unset($data_validation);
$workbook->close();
