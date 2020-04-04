<?php


namespace FFILibXlsxWriter\Types;


use FFI;
use FFI\CData;

abstract class LXWStruct
{
    /**
     * @var CData
     */
    protected CData $struct;

    /**
     * @var CData
     */
    protected CData $pointer;

    public function __destruct()
    {
        $this->free();
    }

    public function getPointer(): CData
    {
        return $this->pointer;
    }

    public function getStruct(): CData
    {
        return $this->struct;
    }

    public function free(): void
    {
        if (!FFI::isNull($this->pointer)) {
            FFI::free($this->pointer);
        }
    }
}
