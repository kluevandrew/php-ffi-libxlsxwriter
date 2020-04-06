<?php

namespace Tests\Structs;

use FFILibXlsxWriter\Exceptions\ZeroDataException;
use FFILibXlsxWriter\Structs\ImageBuffer;
use Tests\FFITestCase;

class ImageBufferTest extends FFITestCase
{
    public function testZeroData(): void
    {
        $this->expectException(ZeroDataException::class);
        $buffer = new ImageBuffer([]);
    }
}
