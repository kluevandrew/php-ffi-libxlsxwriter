<?php

namespace FFILibXlsxWriter\Structs;

use FFI\CData;
use FFILibXlsxWriter\FFILibXlsxWriter;

class ImageBuffer extends Struct
{
    /**
     * @var CData
     */
    protected CData $struct;

    /**
     * ImageBuffer constructor.
     * @param int[] $buffer
     */
    public function __construct(array $buffer)
    {
        $ffi = FFILibXlsxWriter::ffi();
        $type = \FFI::arrayType(\FFI::type('unsigned char'), [count($buffer)]);
        $this->struct = $ffi->new($type, false, false);
        foreach ($buffer as $i => $value) {
            $this->struct[$i] = $value;
        }
        $this->pointer = \FFI::addr($this->struct);
    }
}
