<?php

namespace Tests;

use FFILibXlsxWriter\FFILibXlsxWriter;
use PHPUnit\Framework\TestCase;

class FFITestCase extends TestCase
{
    protected function setUp(): void
    {
        FFILibXlsxWriter::init();
    }
}
