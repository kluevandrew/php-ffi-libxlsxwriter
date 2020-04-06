<?php

namespace Tests\Structs;

use FFILibXlsxWriter\Enums\ValidationCriteria;
use FFILibXlsxWriter\Exceptions\DataValidationException;
use FFILibXlsxWriter\Structs\DataValidation;
use FFILibXlsxWriter\Workbook;
use Tests\FFITestCase;

class DataValidationTest extends FFITestCase
{
    /**
     * @throws \FFILibXlsxWriter\Exceptions\CallAfterCloseException
     * @throws \FFILibXlsxWriter\Exceptions\FFILibXlsxWriterException
     */
    public function testValidationException(): void
    {
        $workbook = new Workbook(sys_get_temp_dir() . '/test.xlsx');
        $worksheet = $workbook->addWorksheet();
        $dataValidation = new DataValidation(ValidationCriteria::NONE);
        $this->expectException(DataValidationException::class);
        $worksheet->dataValidationCell('A1', $dataValidation);
    }
}
