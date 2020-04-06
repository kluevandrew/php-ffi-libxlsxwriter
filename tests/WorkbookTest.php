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
        ];
    }
}
