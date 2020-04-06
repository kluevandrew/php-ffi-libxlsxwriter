<?php

namespace FFILibXlsxWriter\Structs;

use FFILibXlsxWriter\Exceptions\ZeroDataException;
use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Struct;

class ImageBuffer extends Struct
{
    /**
     * ImageBuffer constructor.
     * @param int[] $buffer
     * @throws ZeroDataException
     */
    public function __construct(array $buffer)
    {
        if (empty($buffer)) {
            throw new ZeroDataException(__CLASS__);
        }
        $ffi = FFILibXlsxWriter::ffi();
        $type = \FFI::arrayType(\FFI::type('unsigned char'), [count($buffer)]);
        $this->struct = $ffi->new($type, false, false);
        foreach ($buffer as $i => $value) {
            $this->struct[$i] = $value;
        }
        $this->pointer = \FFI::addr($this->struct);
    }
}
