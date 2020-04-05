<?php

namespace FFILibXlsxWriter\Structs;

use FFI;
use FFILibXlsxWriter\FFILibXlsxWriter;

class CharPointer
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
     * @var FFI\CData
     */
    protected FFI\CData $pointer;

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

    public function __destruct()
    {
        $this->free();
    }

    public function free(): void
    {
        if (!FFI::isNull($this->pointer)) {
            FFI::free($this->pointer);
        }
    }

    private function byteByByteCopy(string $string)
    {
        $ffi = FFILibXlsxWriter::ffi();

        $strlen = strlen($string);
        $this->size = $strlen + 1;
        $this->pointer = $ffi->malloc($this->size * FFI::sizeof(FFI::type('char')));
        $this->pointer = FFI::cast('char*', $this->pointer);

        for ($i = 0; $i < $strlen; $i++) {
            $this->pointer[$i] = $string[$i];
        }

        $this->pointer[$strlen] = "\0";
    }

    /**
     * @return FFI\CData
     */
    public function getPointer(): FFI\CData
    {
        return $this->pointer;
    }

    /**
     * @return string
     */
    public function getString(): string
    {
        return $this->string;
    }
}
