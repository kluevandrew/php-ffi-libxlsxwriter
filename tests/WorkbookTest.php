<?php

namespace Tests;

use FFILibXlsxWriter\Exceptions\CallAfterCloseException;
use FFILibXlsxWriter\Exceptions\FFILibXlsxWriterException;
use FFILibXlsxWriter\Options\DocProperties;
use FFILibXlsxWriter\Structs\DateTime;
use FFILibXlsxWriter\Workbook;

class WorkbookTest extends FFITestCase
{
    /**
     * @param callable $callable
     * @dataProvider callAfterCloseProvider
     * @throws CallAfterCloseException
     */
    public function testCallAfterClose(callable $callable): void
    {
        $workbook = new Workbook(sys_get_temp_dir() . '/test.xlsx');
        $callable($workbook);
        $workbook->close();

        $this->expectException(CallAfterCloseException::class);

        $callable($workbook);
    }

    public function callAfterCloseProvider(): array
    {
        return [
            'addWorksheet' => [fn(Workbook $workbook) => $workbook->addWorksheet()],
            'addFormat' => [fn(Workbook $workbook) => $workbook->addFormat()],
            'getDefaultUrlFormat' => [fn(Workbook $workbook) => $workbook->getDefaultUrlFormat()],
            'addChartsheet' => [fn(Workbook $workbook) => $workbook->addChartsheet()],
//            'addChart' => [fn(Workbook $workbook) => $workbook->addChart()],
            'setProperties' => [fn(Workbook $workbook) => $workbook->setProperties(new DocProperties())],
            'setCustomPropertyString' => [
                fn(Workbook $workbook) => $workbook->setCustomPropertyString('k', 'v')
            ],
            'setCustomPropertyNumber' => [
                fn(Workbook $workbook) => $workbook->setCustomPropertyNumber('n', 1)
            ],
            'setCustomPropertyBoolean' => [
                fn(Workbook $workbook) => $workbook->setCustomPropertyBoolean('b', false)
            ],
            'setCustomPropertyDateTime' => [
                fn(Workbook $workbook) => $workbook->setCustomPropertyDateTime('d', DateTime::now())
            ],
            'defineName' => [fn(Workbook $workbook) => $workbook->defineName('test', '=A1')],
            'getWorksheetByName' => [fn(Workbook $workbook) => $workbook->getWorksheetByName('any')],
            'getChartsheetByName' => [fn(Workbook $workbook) => $workbook->getChartsheetByName('any')],
            'validateSheetname' => [fn(Workbook $workbook) => $workbook->validateSheetname('any')],
//            'addVbaProject' => [fn(Workbook $workbook) => $workbook->addVbaProject()],
            'setVbaName' => [fn(Workbook $workbook) => $workbook->setVbaName('name')],
        ];
    }

    /**
     * @throws CallAfterCloseException
     * @throws FFILibXlsxWriterException
     */
    public function testGetWorksheetByName(): void
    {
        $workbook = new Workbook(sys_get_temp_dir() . '/test.xlsx');

        $this->assertNull($workbook->getWorksheetByName('Sheet1'));
        $this->assertNull($workbook->getWorksheetByName('Sheet2'));
        $this->assertNull($workbook->getWorksheetByName('namedWorksheet'));
        $this->assertNull($workbook->getWorksheetByName('Random name'));

        $unnamedWorksheet = $workbook->addWorksheet();
        $namedWorksheet = $workbook->addWorksheet('namedWorksheet');

        $this->assertEquals($unnamedWorksheet, $workbook->getWorksheetByName('Sheet1'));
        $this->assertNull($workbook->getWorksheetByName('Sheet2'));
        $this->assertEquals($namedWorksheet, $workbook->getWorksheetByName('namedWorksheet'));
        $this->assertNull($workbook->getWorksheetByName('Random name'));

        $this->assertNull($workbook->getChartsheetByName('Sheet1'));
        $this->assertNull($workbook->getChartsheetByName('Sheet12'));
        $this->assertNull($workbook->getChartsheetByName('namedWorksheet'));
        $this->assertNull($workbook->getChartsheetByName('Random name'));
    }

    /**
     * @throws CallAfterCloseException
     * @throws FFILibXlsxWriterException
     */
    public function testGetChartsheetByName(): void
    {
        $workbook = new Workbook(sys_get_temp_dir() . '/test.xlsx');

        $this->assertNull($workbook->getChartsheetByName('Chart1'));
        $this->assertNull($workbook->getChartsheetByName('Chart2'));
        $this->assertNull($workbook->getChartsheetByName('namedChartsheet'));
        $this->assertNull($workbook->getChartsheetByName('Random name'));

        $unnamedChartsheet = $workbook->addChartsheet();
        $namedChartsheet = $workbook->addChartsheet('namedChartsheet');

        $this->assertEquals($unnamedChartsheet, $workbook->getChartsheetByName('Chart1'));
        $this->assertNull($workbook->getChartsheetByName('Chart2'));
        $this->assertEquals($namedChartsheet, $workbook->getChartsheetByName('namedChartsheet'));
        $this->assertNull($workbook->getChartsheetByName('Random name'));

        $this->assertNull($workbook->getWorksheetByName('Chart1'));
        $this->assertNull($workbook->getWorksheetByName('Chart2'));
        $this->assertNull($workbook->getWorksheetByName('namedChartsheet'));
        $this->assertNull($workbook->getWorksheetByName('Random name'));
    }

    /**
     * @throws CallAfterCloseException
     * @throws FFILibXlsxWriterException
     */
    public function testGetByNameUTF8(): void
    {
        $workbook = new Workbook(sys_get_temp_dir() . '/test.xlsx');

        $this->assertNull($workbook->getWorksheetByName('Привет'));
        $this->assertNull($workbook->getChartsheetByName('Мир'));
        $this->assertNull($workbook->getWorksheetByName('Как'));
        $this->assertNull($workbook->getChartsheetByName('Дела'));

        $worksheet = $workbook->addWorksheet('Привет');
        $chartsheet = $workbook->addChartsheet('Мир');

        $this->assertEquals($worksheet, $workbook->getWorksheetByName('Привет'));
        $this->assertEquals($chartsheet, $workbook->getChartsheetByName('Мир'));
        $this->assertNull($workbook->getWorksheetByName('Как'));
        $this->assertNull($workbook->getChartsheetByName('Дела'));
    }
}
