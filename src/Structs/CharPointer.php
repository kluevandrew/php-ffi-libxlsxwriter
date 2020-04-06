<?php

namespace FFILibXlsxWriter\Structs;

use FFI;
use FFILibXlsxWriter\FFILibXlsxWriter;
use FFILibXlsxWriter\Struct;

class CharPointer extends Struct
{
    /**
     * @var string
     */
    protected string $string;

    /**
     * @var int
     */
    protected int $size;

    /**
     * CharPointer constructor.
     * @param string $string
     */
    public function __construct(string $string)
    {
        $this->string = $string;
        $this->size = strlen($string);
        $this->byteByByteCopy($string);
    }

    private function byteByByteCopy(string $string)
    {
        $ffi = FFILibXlsxWriter::ffi();

        $strlen = strlen($string);

        // char* should be a pointer to a Null-terminated string
        $this->size = $strlen + 1;

        $this->struct = $ffi->new(FFI::arrayType(FFI::type('char'), [$this->size]), false, false);
        $this->pointer = FFI::addr($this->struct);

        for ($i = 0; $i < $strlen; $i++) {
            $this->struct[$i] = $string[$i];
        }

        $this->struct[$strlen] = "\0";
    }

    /**
     * @return string
     */
    public function __toString(): string
    {
        return $this->string;
    }
}
