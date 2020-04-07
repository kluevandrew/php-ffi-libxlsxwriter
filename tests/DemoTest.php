<?php

namespace Tests;

use PHPUnit\Framework\TestCase;

class DemoTest extends TestCase
{
    /**
     * @dataProvider demosDataProvider
     */
    public function testAllDemos(string $file): void
    {
        $basename = basename($file, '.php');
//        echo $basename . PHP_EOL;

        $xlsx = __DIR__ . '/../demos/output/' . $basename . '.xlsx';
        if ($basename === 'macro') {
            $xlsx = str_replace('.xlsx', '.xlsm', $xlsx);
        }
        if (file_exists($xlsx)) {
            unlink($xlsx);
        }

        include $file;

        $this->assertTrue(file_exists($xlsx));
        $this->assertGreaterThan(0, filesize($xlsx));
    }

    public function demosDataProvider(): array
    {
        return array_map(
            fn($e) => [$e],
            glob(__DIR__ . '/../demos/*.php'),
        );
    }
}
