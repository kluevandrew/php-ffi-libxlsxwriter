<?php

namespace Tests;

use FFILibXlsxWriter\Enums\Color;
use FFILibXlsxWriter\Enums\ValidationCriteria;
use FFILibXlsxWriter\Exceptions\CallAfterCloseException;
use FFILibXlsxWriter\Structs\DataValidation;
use FFILibXlsxWriter\Structs\DateTime;
use FFILibXlsxWriter\Structs\ImageBuffer;
use FFILibXlsxWriter\Structs\RichString;
use FFILibXlsxWriter\Workbook;
use FFILibXlsxWriter\Worksheet;

class WorksheetTest extends FFITestCase
{
    /**
     * @param callable $callable
     * @dataProvider callAfterCloseProvider
     * @throws CallAfterCloseException
     */
    public function testCallAfterClose(callable $callable): void
    {
        $workbook = new Workbook(sys_get_temp_dir() . '/test.xlsx');
        $worksheet = $workbook->addWorksheet();
        $callable($worksheet);
        $workbook->close();

        $this->expectException(CallAfterCloseException::class);

        $callable($worksheet);
    }

    public function callAfterCloseProvider(): array
    {
        return [
            'autofilter' => [fn(Worksheet $worksheet) => $worksheet->autofilter('A1:B10')],
            'dataValidationCell' => [
                fn(Worksheet $worksheet) => $worksheet->dataValidationCell(
                    'A1',
                    new DataValidation(ValidationCriteria::GREATER_THAN)
                )
            ],
            'freezePanes' => [fn(Worksheet $worksheet) => $worksheet->freezePanes('C100')],
            'insertImage' => [
                fn(Worksheet $worksheet) => $worksheet->insertImage('A1', __DIR__ . '/../demos/logo.png')
            ],
            'writeNumber' => [fn(Worksheet $worksheet) => $worksheet->writeNumber('A1', 1)],
            'writeString' => [fn(Worksheet $worksheet) => $worksheet->writeString('A1', 'string')],
            'writeFormula' => [fn(Worksheet $worksheet) => $worksheet->writeFormula('A1', '=A2')],
            'writeArrayFormula' => [
                fn(Worksheet $worksheet) => $worksheet->writeArrayFormula('A1:B2', '{=SUM(B1:C1*B2:C2)}')
            ],
            'writeDatetime' => [fn(Worksheet $worksheet) => $worksheet->writeDatetime('A1', DateTime::now())],
            'writeUrl' => [fn(Worksheet $worksheet) => $worksheet->writeUrl('A1', 'https://google.com')],
            'writeBoolean' => [fn(Worksheet $worksheet) => $worksheet->writeBoolean('A1', true)],
            'writeRichString' => [
                fn(Worksheet $worksheet) => $worksheet->writeRichString(
                    'A1',
                    new RichString('hello', 'world')
                )
            ],
            'writeComment' => [fn(Worksheet $worksheet) => $worksheet->writeComment('A1', 'comment')],
            'writeBlank' => [fn(Worksheet $worksheet) => $worksheet->writeBlank('A1')],
            'setColumn' => [fn(Worksheet $worksheet) => $worksheet->setColumn('A:F', 100)],
            'setRow' => [fn(Worksheet $worksheet) => $worksheet->setRow(0, 50)],
            'mergeRange' => [fn(Worksheet $worksheet) => $worksheet->mergeRange('A1:B5', 'content')],
            'protect' => [fn(Worksheet $worksheet) => $worksheet->protect()],
            'setSelection' => [fn(Worksheet $worksheet) => $worksheet->setSelection('A1:B5')],
            'splitPanes' => [fn(Worksheet $worksheet) => $worksheet->splitPanes(0.5, 0.5)],
            'insertImageBuffer' => [
                fn(Worksheet $worksheet) => $worksheet->insertImageBuffer(
                    'A1',
                    new ImageBuffer($this->getImageBufferData()),
                    200
                )
            ],
            'setTabColor' => [fn(Worksheet $worksheet) => $worksheet->setTabColor(Color::RED)],
            'setHeader' => [fn(Worksheet $worksheet) => $worksheet->setHeader('header')],
            'setFooter' => [fn(Worksheet $worksheet) => $worksheet->setFooter('footer')],
            'setHorizontalBreaks' => [fn(Worksheet $worksheet) => $worksheet->setHorizontalBreaks([20, 0])],
            'setVerticalBreaks' => [fn(Worksheet $worksheet) => $worksheet->setVerticalBreaks([20, 0])],
        ];
    }

    private function getImageBufferData(): array
    {
        return [
            0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d,
            0x49, 0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x20, 0x00, 0x00, 0x00, 0x20,
            0x08, 0x02, 0x00, 0x00, 0x00, 0xfc, 0x18, 0xed, 0xa3, 0x00, 0x00, 0x00,
            0x01, 0x73, 0x52, 0x47, 0x42, 0x00, 0xae, 0xce, 0x1c, 0xe9, 0x00, 0x00,
            0x00, 0x04, 0x67, 0x41, 0x4d, 0x41, 0x00, 0x00, 0xb1, 0x8f, 0x0b, 0xfc,
            0x61, 0x05, 0x00, 0x00, 0x00, 0x20, 0x63, 0x48, 0x52, 0x4d, 0x00, 0x00,
            0x7a, 0x26, 0x00, 0x00, 0x80, 0x84, 0x00, 0x00, 0xfa, 0x00, 0x00, 0x00,
            0x80, 0xe8, 0x00, 0x00, 0x75, 0x30, 0x00, 0x00, 0xea, 0x60, 0x00, 0x00,
            0x3a, 0x98, 0x00, 0x00, 0x17, 0x70, 0x9c, 0xba, 0x51, 0x3c, 0x00, 0x00,
            0x00, 0x46, 0x49, 0x44, 0x41, 0x54, 0x48, 0x4b, 0x63, 0xfc, 0xcf, 0x40,
            0x63, 0x00, 0xb4, 0x80, 0xa6, 0x88, 0xb6, 0xa6, 0x83, 0x82, 0x87, 0xa6,
            0xce, 0x1f, 0xb5, 0x80, 0x98, 0xe0, 0x1d, 0x8d, 0x03, 0x82, 0xa1, 0x34,
            0x1a, 0x44, 0xa3, 0x41, 0x44, 0x30, 0x04, 0x08, 0x2a, 0x18, 0x4d, 0x45,
            0xa3, 0x41, 0x44, 0x30, 0x04, 0x08, 0x2a, 0x18, 0x4d, 0x45, 0xa3, 0x41,
            0x44, 0x30, 0x04, 0x08, 0x2a, 0x18, 0x4d, 0x45, 0x03, 0x1f, 0x44, 0x00,
            0xaa, 0x35, 0xdd, 0x4e, 0xe6, 0xd5, 0xa1, 0x22, 0x00, 0x00, 0x00, 0x00,
            0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82
        ];
    }
}
