<?php

namespace Tests;

use FFILibXlsxWriter\Exceptions\CallAfterCloseException;
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
            'addWorksheet' => [fn (Workbook $workbook) => $workbook->addWorksheet()],
            'addFormat' => [fn (Workbook $workbook) => $workbook->addFormat()],
            'getDefaultUrlFormat' => [fn (Workbook $workbook) => $workbook->getDefaultUrlFormat()],
//            'addChartsheet' => [fn(Workbook $workbook) => $workbook->addChartsheet()],
//            'addChart' => [fn(Workbook $workbook) => $workbook->addChart()],
//            'setProperties' => [fn(Workbook $workbook) => $workbook->setProperties()],
//            'setCustomPropertyString' => [fn(Workbook $workbook) => $workbook->setCustomPropertyString()],
//            'setCustomPropertyNumber' => [fn(Workbook $workbook) => $workbook->setCustomPropertyNumber()],
//            'setCustomPropertyBoolean' => [fn(Workbook $workbook) => $workbook->setCustomPropertyBoolean()],
//            'setCustomPropertyDateTime' => [fn(Workbook $workbook) => $workbook->setCustomPropertyDateTime()],
//            'defineName' => [fn(Workbook $workbook) => $workbook->defineName()],
//            'getWorksheetByName' => [fn(Workbook $workbook) => $workbook->getWorksheetByName()],
//            'getChartsheetByName' => [fn(Workbook $workbook) => $workbook->getChartsheetByName()],
//            'validateSheetname' => [fn(Workbook $workbook) => $workbook->validateSheetname()],
//            'addVbaProject' => [fn(Workbook $workbook) => $workbook->addVbaProject()],
//            'setVbaName' => [fn(Workbook $workbook) => $workbook->setVbaName()],
        ];
    }
}
